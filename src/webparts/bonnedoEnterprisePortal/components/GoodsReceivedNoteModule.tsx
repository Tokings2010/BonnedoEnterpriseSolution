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

export interface IGoodsReceivedNoteFormProps {
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
    PO_NumberId: number | undefined;
    VendorId: number | undefined;
    Quantity_Received: string;
    Quantity_Accepted: string;
    Quantity_Rejected: string;
    Unit_Price: number | undefined;
    Total_Value: number | undefined;
    Received_Date: string;
    Condition: string;
    Remarks: string;
    Received_ById: number | undefined;
    Inspected_ById: number | undefined;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    projects: IListItem[];
    vendors: IListItem[];
    purchaseOrders: IListItem[];
    isLoadingLookups: boolean;
    currentUserId: number | undefined;
}

const CONDITION_OPTIONS: IDropdownOption[] = [
    { key: 'Good', text: 'Good' },
    { key: 'Damaged', text: 'Damaged' },
    { key: 'Partial', text: 'Partial' },
];

const GoodsReceivedNoteForm: React.FC<IGoodsReceivedNoteFormProps> = ({
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
            PO_NumberId: undefined,
            VendorId: undefined,
            Quantity_Received: '',
            Quantity_Accepted: '',
            Quantity_Rejected: '',
            Unit_Price: undefined,
            Total_Value: undefined,
            Received_Date: '',
            Condition: 'Good',
            Remarks: '',
            Received_ById: undefined,
            Inspected_ById: undefined,
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        projects: [],
        vendors: [],
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
                const [projects, vendors, purchaseOrders, currentUser] = await Promise.all([
                    sharePointService.getProjects(),
                    sharePointService.getVendors(),
                    // Expand Vendor so auto-fill works reliably when selecting a PO
                    sharePointService.getListData('PRC_Purchase_Order_Register', undefined, 100, 0, 'Vendor'),
                    sharePointService.getCurrentUser(),
                ]);

                // Get current user ID
                const userId = currentUser?.ID;

                // Only show Approved POs for GRN creation (STEP 2 enforcement at UI level)
                const approvedPOs = purchaseOrders.filter((po: any) => {
                  const s = (po.Approval_Status || po.Status || '').toLowerCase();
                  return s === 'approved' || s === 'completed';
                });

                setState((prev) => ({
                    ...prev,
                    projects,
                    vendors,
                    purchaseOrders: approvedPOs,
                    isLoadingLookups: false,
                    currentUserId: userId,
                    formData: userId ? {
                        ...prev.formData,
                        Received_ById: userId,
                        Inspected_ById: userId,
                    } : prev.formData,
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

    // Auto-generate GRN Number
    React.useEffect(() => {
        if (isOpen && !state.formData.Title) {
            const generateCode = async (): Promise<void> => {
                try {
                    // Get all items and find the highest number
                    const allItems = await sharePointService.getListData('PRC_GRN_Register', undefined, 500, 0, undefined);
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
                    const newCode = `GRN-${nextNumber.toString().padStart(4, '0')}`;
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Title: newCode },
                    }));
                } catch (error) {
                    console.error('Error generating GRN code:', error);
                    // Fallback to timestamp-based code
                    const timestamp = Date.now().toString().slice(-6);
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Title: `GRN-${timestamp}` },
                    }));
                }
            };
            generateCode();
        }
    }, [isOpen, sharePointService]);

    // Calculate total value when quantity or unit price changes
    React.useEffect(() => {
        const quantity = parseFloat(state.formData.Quantity_Accepted) || 0;
        const unitPrice = state.formData.Unit_Price || 0;
        const total = quantity * unitPrice;

        if (total > 0 && state.formData.Total_Value !== total) {
            setState((prev) => ({
                ...prev,
                formData: { ...prev.formData, Total_Value: total },
            }));
        }
    }, [state.formData.Quantity_Accepted, state.formData.Unit_Price]);

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

    // Auto-fill from selected Purchase Order
    const handlePOChange = (poId: number | undefined): void => {
        if (!poId) {
            handleInputChange('PO_NumberId', undefined);
            return;
        }

        const selectedPO = state.purchaseOrders.find(po => po.Id === poId);
        if (!selectedPO) {
            handleInputChange('PO_NumberId', poId);
            return;
        }

        const updates: Partial<IFormData> = {
            PO_NumberId: poId,
        };

        // Support both direct ID and expanded lookup object
        const vendorId = selectedPO.VendorId || (selectedPO.Vendor && (selectedPO.Vendor.Id || selectedPO.Vendor.ID));
        if (vendorId) {
            updates.VendorId = vendorId;
        }

        if (selectedPO.Quantity) {
            updates.Quantity_Received = String(selectedPO.Quantity);
            updates.Quantity_Accepted = String(selectedPO.Quantity);
        }

        if (selectedPO.UnitPrice) {
            updates.Unit_Price = selectedPO.UnitPrice;
        }

        if (selectedPO.TotalAmount) {
            updates.Total_Value = selectedPO.TotalAmount;
        }

        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                ...updates,
            },
        }));
    };

    const validateForm = (): boolean => {
        if (!state.formData.Project_CodeId) {
            setState((prev) => ({ ...prev, error: 'Project Code is required' }));
            return false;
        }
        if (!state.formData.PO_NumberId) {
            setState((prev) => ({ ...prev, error: 'PO Number is required' }));
            return false;
        }
        if (!state.formData.Quantity_Received.trim()) {
            setState((prev) => ({ ...prev, error: 'Quantity Received is required' }));
            return false;
        }
        if (!state.formData.Received_ById) {
            setState((prev) => ({ ...prev, error: 'Received By is required' }));
            return false;
        }
        return true;
    };

    const handleSubmit = React.useCallback(async () => {
        if (!validateForm()) {
            return;
        }

        // STEP 2: Validation - Cannot create GRN unless PO is Approved
        if (state.formData.PO_NumberId) {
            const selectedPO = state.purchaseOrders.find(po => po.Id === state.formData.PO_NumberId);
            const poStatus = selectedPO?.Approval_Status || selectedPO?.Status;

            if (poStatus && poStatus !== 'Approved') {
                setState((prev) => ({
                    ...prev,
                    error: 'Cannot create GRN: Linked Purchase Order must be Approved first.',
                }));
                return;
            }
        }

        setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

        try {
            const itemData = {
                Title: state.formData.Title || `GRN-${Date.now()}`,
                // NOTE: Project_CodeId and Status removed — they do not exist on PRC_GRN_Register
                PO_NumberId: state.formData.PO_NumberId,
                VendorId: state.formData.VendorId,
                Quantity_Received: parseInt(state.formData.Quantity_Received, 10),
                Quantity_Accepted: parseInt(state.formData.Quantity_Accepted || '0', 10),
                Quantity_Rejected: parseInt(state.formData.Quantity_Rejected || '0', 10),
                Unit_Price: state.formData.Unit_Price || 0,
                Total_Value: state.formData.Total_Value || 0,
                Received_Date: state.formData.Received_Date ? new Date(state.formData.Received_Date).toISOString() : new Date().toISOString(),
                Condition: state.formData.Condition,
                Remarks: state.formData.Remarks,
                Received_ById: state.formData.Received_ById,
                Inspected_ById: state.formData.Inspected_ById,
            };

            /* 
             * Recommended / Safe columns for PRC_GRN_Register (based on current schema + approval needs):
             * - Title, GRN_Number
             * - PO_NumberId (Link to PO)
             * - VendorId
             * - Quantity_Received, Quantity_Accepted, Quantity_Rejected
             * - Unit_Price, Total_Value
             * - Received_Date, Condition, Remarks
             * - Received_ById, Inspected_ById
             *
             * Approval-related (add these for full tracker support, like MR/PR/PO):
             * - Approval_Status, Status (if different), Approval_Level, Current_Approver,
             *   Approval_History, Approval_Started_On, Approval_Completed_On
             *
             * Do NOT send: Project_CodeId, Status (unless the column exists), MaterialId (unless added)
             */
            const createdItem = await sharePointService.createListItem('PRC_GRN_Register', itemData);

            if (createdItem && createdItem.ID) {
                // Best-effort post-creation steps — never block the user if approval columns are missing
                try {
                    await sharePointService.initializeGoodsReceivedNoteApproval(createdItem.ID);
                } catch (e) {
                    console.warn('[GRN Form] Approval init skipped (non-fatal):', e);
                }

                try {
                    if (state.formData.PO_NumberId) {
                        await sharePointService.notifyFinanceManagerAfterGRN(createdItem.ID, state.formData.PO_NumberId);
                    }
                } catch (e) {
                    console.warn('[GRN Form] Finance notification skipped (non-fatal):', e);
                }
            }

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                successMessage: 'Goods Received Note created successfully!',
                formData: {
                    Title: '',
                    Project_CodeId: undefined,
                    PO_NumberId: undefined,
                    VendorId: undefined,
                    Quantity_Received: '',
                    Quantity_Accepted: '',
                    Quantity_Rejected: '',
                    Unit_Price: undefined,
                    Total_Value: undefined,
                    Received_Date: '',
                    Condition: 'Good',
                    Remarks: '',
                    Received_ById: undefined,
                    Inspected_ById: undefined,
                },
            }));

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to create Goods Received Note';
            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                error: errorMessage,
            }));
        }
    }, [state.formData, state.purchaseOrders, sharePointService, onDismiss, onSubmitSuccess]);

    const handleCancel = (): void => {
        setState({
            formData: {
                Title: '',
                Project_CodeId: undefined,
                PO_NumberId: undefined,
                VendorId: undefined,
                Quantity_Received: '',
                Quantity_Accepted: '',
                Quantity_Rejected: '',
                Unit_Price: undefined,
                Total_Value: undefined,
                Received_Date: '',
                Condition: 'Good',
                Remarks: '',
                Received_ById: state.currentUserId,
                Inspected_ById: state.currentUserId,
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            projects: state.projects,
            vendors: state.vendors,
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

    const poOptions: IDropdownOption[] = state.purchaseOrders.map((item) => ({
        key: item.Id,
        text: item.Title || item.PO_Number || `PO ${item.Id}`,
    }));

    const isDisabled = state.isSubmitting || state.isLoadingLookups;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText="New Goods Received Note"
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
                            {/* Section 1: Receipt Details */}
                            <div className={classNames.sectionHeader}>Receipt Details</div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="GRN Number"
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
                                    onChange={(e, option) => handleInputChange('Project_CodeId', option?.key as number)}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Project"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Purchase Order"
                                    options={poOptions}
                                    selectedKey={state.formData.PO_NumberId}
                                    onChange={(e, option) => handlePOChange(option?.key as number)}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select PO"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Vendor"
                                    options={vendorOptions}
                                    selectedKey={state.formData.VendorId}
                                    onChange={(e, option) => handleInputChange('VendorId', option?.key as number)}
                                    disabled={isDisabled}
                                    placeholder="Select Vendor (Optional)"
                                />
                            </div>

                            {/* Section 2: Quantity Details */}
                            <div className={classNames.sectionHeader}>Quantity Details</div>

                            <div className={classNames.fieldGroup}>
                                <Stack horizontal tokens={{ childrenGap: 12 }}>
                                    <TextField
                                        label="Quantity Received"
                                        type="number"
                                        value={state.formData.Quantity_Received}
                                        onChange={(_, value) => handleInputChange('Quantity_Received', value)}
                                        disabled={isDisabled}
                                        required
                                        styles={{ root: { flex: 1 } }}
                                    />
                                    <TextField
                                        label="Quantity Accepted"
                                        type="number"
                                        value={state.formData.Quantity_Accepted}
                                        onChange={(_, value) => handleInputChange('Quantity_Accepted', value)}
                                        disabled={isDisabled}
                                        styles={{ root: { flex: 1 } }}
                                    />
                                    <TextField
                                        label="Quantity Rejected"
                                        type="number"
                                        value={state.formData.Quantity_Rejected}
                                        onChange={(_, value) => handleInputChange('Quantity_Rejected', value)}
                                        disabled={isDisabled}
                                        styles={{ root: { flex: 1 } }}
                                    />
                                </Stack>
                            </div>

                            {/* Section 3: Financial Details */}
                            <div className={classNames.sectionHeader}>Financial Details</div>

                            <div className={classNames.fieldGroup}>
                                <Stack horizontal tokens={{ childrenGap: 12 }}>
                                    <TextField
                                        label="Unit Price"
                                        type="number"
                                        value={state.formData.Unit_Price?.toString() || ''}
                                        onChange={(_, value) => handleInputChange('Unit_Price', value ? Number(value) : undefined)}
                                        disabled={isDisabled}
                                        prefix="$"
                                        styles={{ root: { flex: 1 } }}
                                    />
                                    <TextField
                                        label="Total Value"
                                        type="number"
                                        value={state.formData.Total_Value?.toString() || ''}
                                        onChange={(_, value) => handleInputChange('Total_Value', value ? Number(value) : undefined)}
                                        disabled={isDisabled}
                                        prefix="$"
                                        styles={{ root: { flex: 1 } }}
                                    />
                                </Stack>
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Received Date"
                                    type="date"
                                    value={state.formData.Received_Date}
                                    onChange={(_, value) => handleInputChange('Received_Date', value)}
                                    disabled={isDisabled}
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Condition"
                                    options={CONDITION_OPTIONS}
                                    selectedKey={state.formData.Condition}
                                    onChange={(e, option) => handleInputChange('Condition', option?.key as string)}
                                    disabled={isDisabled}
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Remarks"
                                    value={state.formData.Remarks}
                                    onChange={(_, value) => handleInputChange('Remarks', value)}
                                    disabled={isDisabled}
                                    multiline
                                    rows={3}
                                />
                            </div>

                            {/* Section 4: Receiving Personnel */}
                            <div className={classNames.sectionHeader}>Receiving Personnel</div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="grn-received-by-picker"
                                    titleText="Received By"
                                    selectedUsers={state.formData.Received_ById ? [{ id: typeof state.formData.Received_ById === 'string' ? parseInt(state.formData.Received_ById, 10) : state.formData.Received_ById, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                    required={true}
                                    disabled={isDisabled}
                                    spHttpClient={spHttpClient}
                                    pageContext={pageContext}
                                    webPartContext={webPartContext}
                                    onChange={(items) => {
                                        if (items && items.length > 0 && items[0].id) {
                                            handleInputChange('Received_ById', items[0].id);
                                        } else {
                                            handleInputChange('Received_ById', undefined);
                                        }
                                    }}
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="grn-inspected-by-picker"
                                    titleText="Inspected By"
                                    selectedUsers={state.formData.Inspected_ById ? [{ id: typeof state.formData.Inspected_ById === 'string' ? parseInt(state.formData.Inspected_ById, 10) : state.formData.Inspected_ById, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                    disabled={isDisabled}
                                    spHttpClient={spHttpClient}
                                    pageContext={pageContext}
                                    webPartContext={webPartContext}
                                    onChange={(items) => {
                                        if (items && items.length > 0 && items[0].id) {
                                            handleInputChange('Inspected_ById', items[0].id);
                                        } else {
                                            handleInputChange('Inspected_ById', undefined);
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

export default GoodsReceivedNoteForm;
