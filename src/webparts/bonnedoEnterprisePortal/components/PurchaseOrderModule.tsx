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
    Text,
} from '@fluentui/react';
import PeoplePicker from './PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService, IListItem } from '../services/SharePointService';
import { NotificationService } from '../services/NotificationService';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IPurchaseOrderFormProps {
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
    // Single PR kept for backward compatibility with existing records
    PR_NumberId: number | undefined;
    // New: Supports multiple PRs per PO (PR_References multi-lookup on the list)
    PR_ReferenceIds: number[];
    VendorId: number | undefined;
    Quantity: string;
    UnitPrice: number | undefined;
    TotalAmount: number | undefined;
    Amount: number | undefined;
    Delivery_Date: string;
    Description: string;
    Requested_ById: number | undefined;
    Project_ManagerId: number | undefined;
    Current_ApproverId: number | undefined;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    projects: IListItem[];
    vendors: IListItem[];
    purchaseRequisitions: IListItem[];
    isLoadingLookups: boolean;
    currentUserId: number | undefined;
}

const PurchaseOrderForm: React.FC<IPurchaseOrderFormProps> = ({
    isOpen,
    onDismiss,
    onSubmitSuccess,
    spHttpClient,
    pageContext,
    webPartContext,
}) => {
    const theme = getTheme();

    // Local state for PR summary expansion (auto-collapse when many PRs selected)
    const [showPRSummaryDetails, setShowPRSummaryDetails] = React.useState(true);

    const [state, setState] = React.useState<IFormState>({
        formData: {
            Title: '',
            Project_CodeId: undefined,
            PR_NumberId: undefined,
            PR_ReferenceIds: [],
            VendorId: undefined,
            Quantity: '',
            UnitPrice: undefined,
            TotalAmount: undefined,
            Amount: undefined,
            Delivery_Date: '',
            Description: '',
            Requested_ById: undefined,
            Project_ManagerId: undefined,
            Current_ApproverId: undefined,
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        projects: [],
        vendors: [],
        purchaseRequisitions: [],
        isLoadingLookups: true,
        currentUserId: undefined,
    });

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    // Helper to extract numeric ID from lookup (object or primitive) - matches pattern used in PR/MR forms
    const getLookupId = (val: any): number | undefined => {
        if (typeof val === 'number' || typeof val === 'string') return Number(val) || undefined;
        if (val && typeof val === 'object') return val.ID || val.Id || val.id;
        return undefined;
    };

    // Auto-collapse PR summary when more than 3 PRs are selected (minor polish)
    React.useEffect(() => {
        if (state.formData.PR_ReferenceIds.length > 3) {
            setShowPRSummaryDetails(false);
        } else if (state.formData.PR_ReferenceIds.length <= 3) {
            setShowPRSummaryDetails(true);
        }
    }, [state.formData.PR_ReferenceIds.length]);

    // Load lookup data on mount
    React.useEffect(() => {
        const loadLookups = async (): Promise<void> => {
            try {
                const [projects, vendors, purchaseRequisitionsRaw, currentUser] = await Promise.all([
                    sharePointService.getProjects(),
                    sharePointService.getVendors(),
                    sharePointService.getPurchaseRequisitions(),
                    sharePointService.getCurrentUser(),
                ]);

                // Best-effort: load used PRs for duplicate prevention. 
                // If this fails (e.g. column PR_NumberId or expansion issue), the form should still work.
                let existingPOs: any[] = [];
                try {
                    existingPOs = await sharePointService.getListData(
                        'PRC_Purchase_Order_Register', 
                        undefined, 
                        500, 
                        0, 
                        'PR_References'
                    );
                } catch (e) {
                    console.warn('[PO Form] Could not load existing POs for duplicate prevention (non-fatal):', e);
                }

                // Compute set of PR IDs that already have a PO (supports both legacy single + new multi-lookup PR_References)
                const usedPRIds = new Set<number>();
                (existingPOs || []).forEach((po: any) => {
                    // Legacy single reference (safe — just read the value)
                    const legacy = getLookupId(po.PR_NumberId || po.PR_Number || po.PR_ReferenceId);
                    if (legacy) usedPRIds.add(legacy);

                    // New multi-lookup PR_References
                    const refs = po.PR_References || po.PR_ReferencesId || [];
                    if (Array.isArray(refs)) {
                        refs.forEach((r: any) => {
                            const id = getLookupId(r);
                            if (id) usedPRIds.add(id);
                        });
                    }
                });

                // Filter to only PRs without existing PO (duplicate-safe)
                const purchaseRequisitions = purchaseRequisitionsRaw.filter((pr: any) => {
                    const prId = pr.Id || pr.ID || getLookupId(pr);
                    return prId != null && !usedPRIds.has(prId);
                });

                // Get current user ID
                const userId = currentUser?.ID;

                setState((prev) => ({
                    ...prev,
                    projects,
                    vendors,
                    purchaseRequisitions,
                    isLoadingLookups: false,
                    currentUserId: userId,
                    formData: userId ? {
                        ...prev.formData,
                        Requested_ById: userId,
                        Current_ApproverId: userId,
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

    // Auto-generate PO Number and calculate total
    React.useEffect(() => {
        if (isOpen && !state.formData.Title) {
            const generateCode = async (): Promise<void> => {
                try {
                    // Get all items and find the highest number
                    const allItems = await sharePointService.getListData('PRC_Purchase_Order_Register', undefined, 500, 0, undefined);
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
                    const newCode = `BON-PO-${nextNumber.toString().padStart(4, '0')}`;
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Title: newCode },
                    }));
                } catch (error) {
                    console.error('Error generating PO code:', error);
                    // Fallback to timestamp-based code
                    const timestamp = Date.now().toString().slice(-6);
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Title: `PO-${timestamp}` },
                    }));
                }
            };
            generateCode();
        }
    }, [isOpen, sharePointService]);

    // Calculate total when quantity or unit price changes
    React.useEffect(() => {
        const quantity = parseFloat(state.formData.Quantity) || 0;
        const unitPrice = state.formData.UnitPrice || 0;
        const total = quantity * unitPrice;

        if (total > 0 && state.formData.TotalAmount !== total) {
            setState((prev) => ({
                ...prev,
                formData: { ...prev.formData, TotalAmount: total },
            }));
        }
    }, [state.formData.Quantity, state.formData.UnitPrice]);

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

    // Auto-populate PO form fields when a Purchase Requisition is selected (replicates MR->PR prefill behavior + duplicate safety)
    const handlePurchaseRequisitionChange = (prId: number | undefined): void => {
        handleInputChange('PR_NumberId', prId);

        if (!prId) {
            return;
        }

        const selectedPR = state.purchaseRequisitions.find((pr: any) => (pr.Id || pr.ID) === prId);
        if (!selectedPR) {
            return;
        }

        // Project (and auto Project Manager like the Project dropdown handler)
        const projId = getLookupId(selectedPR.Project_CodeId) || getLookupId(selectedPR.Project_Code);
        if (projId) {
            handleInputChange('Project_CodeId', projId);
            const selectedProject = state.projects.find((p: any) => (p.Id || p.ID) === projId);
            if (selectedProject && selectedProject.Project_ManagerId) {
                handleInputChange('Project_ManagerId', selectedProject.Project_ManagerId);
            }
        }

        // Description (UI only)
        if (selectedPR.Description || selectedPR.Title) {
            handleInputChange('Description', selectedPR.Description || selectedPR.Title || '');
        }

        // Quantity from PR (for calculation and display)
        const prQty = selectedPR.Quantity != null ? Number(selectedPR.Quantity) : 0;
        if (prQty > 0) {
            handleInputChange('Quantity', String(prQty));
        }

        // Monetary value from PR becomes the authoritative Total/Amount (matches Convert-to-PO behaviour)
        const estCost = selectedPR.Estimated_Cost ?? selectedPR.EstimatedCost ?? selectedPR.TotalAmount ?? selectedPR.Amount;
        if (estCost !== undefined && estCost !== null) {
            const numCost = Number(estCost) || 0;
            handleInputChange('TotalAmount', numCost);
            handleInputChange('Amount', numCost);

            // Prefill Unit Price by back-calculation when we have Quantity (UX only – not saved to list)
            if (prQty > 0 && numCost > 0) {
                handleInputChange('UnitPrice', numCost / prQty);
            }
        }

        // Delivery Date (normalize to YYYY-MM-DD for date input)
        const del = selectedPR.Delivery_Date || selectedPR.Required_Date || selectedPR.DeliveryDate || selectedPR.RequiredDate;
        if (del) {
            const dstr = typeof del === 'string' ? del.substring(0, 10) : new Date(del).toISOString().substring(0, 10);
            handleInputChange('Delivery_Date', dstr);
        }

        // Requested By (if PR carries it)
        const reqBy = getLookupId(selectedPR.Requested_ById) || getLookupId(selectedPR.Requested_By);
        if (reqBy) {
            handleInputChange('Requested_ById', reqBy);
        }
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
        if (!state.formData.Quantity.trim()) {
            setState((prev) => ({ ...prev, error: 'Quantity is required' }));
            return false;
        }
        // UnitPrice is UI/calculation only for the form (not a column on the PO list).
        // Require either (Quantity + UnitPrice) for manual entry, or a TotalAmount (populated from PR or entered).
        const hasTotal = !!state.formData.TotalAmount && Number(state.formData.TotalAmount) > 0;
        const hasQtyAndPrice = state.formData.Quantity && state.formData.UnitPrice;
        if (!hasTotal && !hasQtyAndPrice) {
            setState((prev) => ({ ...prev, error: 'Either Total Amount or Quantity + Unit Price is required' }));
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
            const totalAmount = state.formData.TotalAmount || (parseFloat(state.formData.Quantity) * (state.formData.UnitPrice || 0));

            const itemData: any = {
                Title: state.formData.Title || `BON-PO-${Date.now()}`,
                PO_Number: state.formData.Title || `BON-PO-${Date.now()}`,
                PR_NumberId: state.formData.PR_NumberId,           // backward compat (first PR)
                Project_CodeId: state.formData.Project_CodeId,
                VendorId: state.formData.VendorId,
                Amount: totalAmount,
                TotalAmount: totalAmount,
                Delivery_Date: state.formData.Delivery_Date ? new Date(state.formData.Delivery_Date).toISOString() : null,
                Status: 'Draft',
            };

            // Multi-PR support via PR_References (multi-lookup)
            // IMPORTANT: The column "PR_References" (Lookup, Allow multiple values = Yes) must be created
            // on PRC_Purchase_Order_Register before we can write {"results": [...]}
            // Until then, we only write the first PR to the legacy PR_NumberId field.
            if (state.formData.PR_ReferenceIds && state.formData.PR_ReferenceIds.length > 0) {
                // Always store the first PR in the legacy single field (safe)
                if (!itemData.PR_NumberId) {
                    itemData.PR_NumberId = state.formData.PR_ReferenceIds[0];
                }

                // Only attempt the multi-lookup write if the user selected more than 1 PR
                // and the column has been created on the list.
                if (state.formData.PR_ReferenceIds.length > 1) {
                    // Temporarily commented out until the PR_References multi-lookup column exists
                    // itemData.PR_ReferencesId = { results: state.formData.PR_ReferenceIds };
                    console.warn(
                        '[PO Form] Multiple PRs selected, but PR_References multi-lookup column does not exist yet. ' +
                        'Only the first PR was saved. Create the column to enable full multi-PR support.'
                    );
                }
            }

            const createdItem = await sharePointService.createListItem('PRC_Purchase_Order_Register', itemData);

            // Initialize approval with threshold logic (unique PO amount-based flow)
            if (createdItem && createdItem.ID) {
                await sharePointService.initializePurchaseOrderApproval(createdItem.ID, totalAmount);

                // Fetch updated record to get Current_Approver set by the amount-threshold logic, then notify
                try {
                  const updatedItems = await sharePointService.getListData('PRC_Purchase_Order_Register', `ID eq ${createdItem.ID}`, 1, 0, 'Project_Code,PR_Number,Vendor');
                  if (updatedItems.length > 0) {
                     const fresh = updatedItems[0];
                     const ca = fresh?.Current_Approver;
                     const approverEmail = ca
                       ? (typeof ca === 'object' ? (ca.Email || ca.Title || ca.LoginName) : ca)
                       : null;

                     if (approverEmail && webPartContext) {
                      const graphClient: MSGraphClientV3 = await webPartContext.msGraphClientFactory.getClient('3');
                      const notificationService = new NotificationService(graphClient);

                      const projectName = fresh.Project_Code?.Project_Code || fresh.Project_Code?.Title || 'Unknown Project';
                      const deepLink = `${window.location.origin}${window.location.pathname}?po=${createdItem.ID}`;

                      console.log('[PurchaseOrderForm] Sending PO approval notification to:', approverEmail);

                      await notificationService.sendMaterialRequestApprovalNotification({
                        requestTitle: fresh.PO_Number || fresh.Title || `PO-${createdItem.ID}`,
                        project: projectName,
                        amount: totalAmount,
                        approverEmail,
                        deepLink,
                      });
                    }
                  }
                } catch (notifyErr) {
                  console.error('[PurchaseOrderForm] Notification after init failed:', notifyErr);
                }
            }

            // Mark ALL selected PRs as Completed (full multi-PR support)
            // This prevents duplicate POs for any of the linked PRs and replicates the Convert behavior for all of them.
            const prsToMark = (state.formData.PR_ReferenceIds && state.formData.PR_ReferenceIds.length > 0)
                ? state.formData.PR_ReferenceIds
                : (state.formData.PR_NumberId ? [state.formData.PR_NumberId] : []);

            for (const prId of prsToMark) {
                try {
                    await sharePointService.updateListItem('PRC_Purchase_Requisition_Register', prId, {
                        Approval_Status: 'Approved',
                        Status: 'Completed',
                    });
                } catch (markErr) {
                    console.warn('[PurchaseOrderForm] Could not mark PR as completed:', prId, markErr);
                }
            }

            // Trigger PDF generation with all linked PRs so the document includes line items from every selected PR
            if (webPartContext && state.formData.Title) {
                try {
                    await sharePointService.triggerPurchaseOrderPdfGeneration(
                        state.formData.Title,
                        createdItem.ID,
                        prsToMark
                    );
                } catch (pdfErr) {
                    console.warn('[PurchaseOrderForm] PDF trigger failed (non-blocking):', pdfErr);
                }
            }

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                successMessage: 'Purchase Order created successfully!',
                formData: {
                    Title: '',
                    Project_CodeId: undefined,
                    PR_NumberId: undefined,
                    PR_ReferenceIds: [],
                    VendorId: undefined,
                    Quantity: '',
                    UnitPrice: undefined,
                    TotalAmount: undefined,
                    Amount: undefined,
                    Delivery_Date: '',
                    Description: '',
                    Requested_ById: undefined,
                    Project_ManagerId: undefined,
                    Current_ApproverId: undefined,
                },
            }));

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to create Purchase Order';
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
                Title: '',
                Project_CodeId: undefined,
                PR_NumberId: undefined,
                PR_ReferenceIds: [],
                VendorId: undefined,
                Quantity: '',
                UnitPrice: undefined,
                TotalAmount: undefined,
                Amount: undefined,
                Delivery_Date: '',
                Description: '',
                Requested_ById: undefined,
                Project_ManagerId: undefined,
                Current_ApproverId: undefined,
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            projects: state.projects,
            vendors: state.vendors,
            purchaseRequisitions: state.purchaseRequisitions,
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
        text: item.Title || item.PR_Number || `PR ${item.Id}`,
    }));

    const isDisabled = state.isSubmitting || state.isLoadingLookups;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText="New Purchase Order"
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
                            {/* Section 1: Order Details */}
                            <div className={classNames.sectionHeader}>Order Details</div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="PO Number"
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
                                        const selectedProject = state.projects.find(p => p.Id === option?.key);
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
                                    label="Purchase Requisition Reference(s)"
                                    options={prOptions}
                                    selectedKeys={state.formData.PR_ReferenceIds}
                                    multiSelect
                                    onChange={(e, option) => {
                                        if (!option) return;
                                        const current = [...(state.formData.PR_ReferenceIds || [])];
                                        const key = Number(option.key);
                                        const idx = current.indexOf(key);
                                        if (idx > -1) {
                                            current.splice(idx, 1);
                                        } else {
                                            current.push(key);
                                        }
                                        // Update array
                                        setState(prev => ({
                                            ...prev,
                                            formData: { ...prev.formData, PR_ReferenceIds: current }
                                        }));
                                        // For compatibility, set first as PR_NumberId and trigger rich prefill from first
                                        if (current.length > 0) {
                                            handlePurchaseRequisitionChange(current[0]);
                                        } else {
                                            handleInputChange('PR_NumberId', undefined);
                                        }
                                    }}
                                    disabled={isDisabled}
                                    placeholder="Select one or more PRs (Multi-select supported)"
                                />
                                <Text variant="small" style={{ color: theme.palette.neutralSecondary, marginTop: 4 }}>
                                    You can select multiple PRs — they will all be included in this PO and the generated PDF.
                                </Text>

                                {/* Enhanced polish: Selected PRs summary with auto-collapse, total value, and project warning */}
                                {state.formData.PR_ReferenceIds.length > 0 && (() => {
                                    const selectedIds = state.formData.PR_ReferenceIds;
                                    const selectedPRs = state.purchaseRequisitions.filter((pr: any) =>
                                        selectedIds.includes(pr.Id || pr.ID)
                                    );

                                    // Total estimated value across all selected PRs
                                    const totalEstValue = selectedPRs.reduce((sum: number, pr: any) => {
                                        const val = Number(pr.Estimated_Cost ?? pr.EstimatedCost ?? pr.TotalAmount ?? pr.Amount ?? 0);
                                        return sum + (isNaN(val) ? 0 : val);
                                    }, 0);

                                    // Project mismatch detection for warning
                                    const projectIds = new Set(
                                        selectedPRs.map((pr: any) =>
                                            getLookupId(pr.Project_CodeId) || getLookupId(pr.Project_Code)
                                        ).filter(Boolean)
                                    );
                                    const hasProjectConflict = projectIds.size > 1;

                                    const isCollapsed = !showPRSummaryDetails && selectedIds.length > 3;
                                    const displayCount = selectedIds.length;

                                    return (
                                        <div style={{ marginTop: 6, padding: '8px 12px', backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4, border: hasProjectConflict ? `1px solid ${theme.palette.yellow}` : 'none' }}>
                                            {/* Header row */}
                                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
                                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                                    <Text variant="small" style={{ fontWeight: 600 }}>
                                                        Selected PRs ({displayCount})
                                                    </Text>
                                                    {totalEstValue > 0 && (
                                                        <Text variant="small" style={{ color: theme.palette.themePrimary, fontWeight: 600 }}>
                                                            Total Est: ${totalEstValue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                                                        </Text>
                                                    )}
                                                </div>

                                                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                                                    {selectedIds.length > 3 && (
                                                        <DefaultButton
                                                            text={showPRSummaryDetails ? 'Collapse' : 'Show details'}
                                                            onClick={() => setShowPRSummaryDetails(!showPRSummaryDetails)}
                                                            styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 24, fontSize: 12 } }}
                                                        />
                                                    )}
                                                    <DefaultButton
                                                        text="Clear all"
                                                        onClick={() => {
                                                            setShowPRSummaryDetails(true);
                                                            setState(prev => ({
                                                                ...prev,
                                                                formData: { ...prev.formData, PR_ReferenceIds: [], PR_NumberId: undefined }
                                                            }));
                                                        }}
                                                        styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 24, fontSize: 12 } }}
                                                    />
                                                </div>
                                            </div>

                                            {/* Project conflict warning */}
                                            {hasProjectConflict && (
                                                <div style={{ marginBottom: 6 }}>
                                                    <MessageBar messageBarType={MessageBarType.warning} isMultiline={false} styles={{ root: { padding: '4px 8px' } }}>
                                                        Selected PRs belong to different projects. The PO will use the project from the first selected PR.
                                                    </MessageBar>
                                                </div>
                                            )}

                                            {/* Details (pills) - auto-collapsed when > 3 */}
                                            {!isCollapsed && (
                                                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                                                    {selectedIds.map(id => {
                                                        const pr = state.purchaseRequisitions.find((p: any) => (p.Id || p.ID) === id);
                                                        const label = pr ? (pr.Title || pr.PR_Number || `PR ${id}`) : `PR ${id}`;
                                                        return (
                                                            <span key={id} style={{
                                                                backgroundColor: theme.palette.themeLighter,
                                                                color: theme.palette.themeDark,
                                                                padding: '2px 8px',
                                                                borderRadius: 12,
                                                                fontSize: 12,
                                                                whiteSpace: 'nowrap'
                                                            }}>
                                                                {label}
                                                            </span>
                                                        );
                                                    })}
                                                </div>
                                            )}

                                            {isCollapsed && (
                                                <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                                                    {displayCount} PRs selected (details hidden)
                                                </Text>
                                            )}
                                        </div>
                                    );
                                })()}
                            </div>

                            {/* Section 2: Vendor Details */}
                            <div className={classNames.sectionHeader}>Vendor Details</div>

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

                            {/* Section 3: Financial Details */}
                            <div className={classNames.sectionHeader}>Financial Details</div>

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
                                    <TextField
                                        label="Unit Price"
                                        type="number"
                                        value={state.formData.UnitPrice?.toString() || ''}
                                        onChange={(_, value) => handleInputChange('UnitPrice', value ? Number(value) : undefined)}
                                        disabled={isDisabled}
                                        required
                                        prefix="$"
                                        styles={{ root: { flex: 1 } }}
                                    />
                                </Stack>
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Total Amount"
                                    type="number"
                                    value={state.formData.TotalAmount?.toString() || ''}
                                    onChange={(_, value) => handleInputChange('TotalAmount', value ? Number(value) : undefined)}
                                    disabled={isDisabled}
                                    prefix="$"
                                    styles={{ root: { backgroundColor: theme.palette.neutralLighterAlt } }}
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Delivery Date"
                                    type="date"
                                    value={state.formData.Delivery_Date}
                                    onChange={(_, value) => handleInputChange('Delivery_Date', value)}
                                    disabled={isDisabled}
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

                            {/* Section 4: Personnel */}
                            <div className={classNames.sectionHeader}>Personnel</div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="po-requested-by-picker"
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
                                    key="po-project-manager-picker"
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

                             {/* Current Approver is now system-determined by the Approval Matrix (threshold + CFO escalation).
                                 The picker was removed to avoid confusion with the automatic approval routing. */}

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

export default PurchaseOrderForm;
