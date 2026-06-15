import * as React from 'react';
import {
    Panel,
    PanelType,
    Text,
    Label,
    IconButton,
    Spinner,
    SpinnerSize,
    getTheme,
    mergeStyleSets,
    ScrollablePane,
    ScrollbarVisibility,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';

export interface IApprovalTrackerPanelProps {
    isOpen: boolean;
    recordId: number;
    listName: string;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onDismiss: () => void;
    onRefresh?: () => void;
}

interface IApprovalTrackerPanelState {
    record: IListItem | undefined;
    isLoading: boolean;
    error: string | undefined;
    isMobileView: boolean;
}

const ApprovalTrackerPanel: React.FC<IApprovalTrackerPanelProps> = ({
    isOpen,
    recordId,
    listName,
    spHttpClient,
    pageContext,
    onDismiss,
    onRefresh,
}) => {
    const theme = getTheme();
    const [state, setState] = React.useState<IApprovalTrackerPanelState>({
        record: undefined,
        isLoading: true,
        error: undefined,
        isMobileView: typeof window !== 'undefined' ? window.innerWidth < 768 : false,
    });

    // Handle window resize for responsive behavior
    React.useEffect(() => {
        const handleResize = (): void => {
            setState((prev) => ({
                ...prev,
                isMobileView: window.innerWidth < 768,
            }));
        };

        window.addEventListener('resize', handleResize);
        return () => window.removeEventListener('resize', handleResize);
    }, []);

    // Fetch record data when panel opens or recordId changes
    React.useEffect(() => {
        if (isOpen && recordId) {
            fetchRecord();
        }
    }, [isOpen, recordId, listName]);

    const fetchRecord = async (): Promise<void> => {
        setState((prev) => ({ ...prev, isLoading: true, error: undefined }));

        try {
            // Safe, minimal query first (many PO lists do not have the full MR/PR approval schema).
            // We deliberately avoid fields that do not exist on PRC_Purchase_Order_Register.
            const webUrl = pageContext.web.absoluteUrl;
            const base = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items(${recordId})`;

            // Conservative select/expand that works for the current PO list schema + still pulls approval fields when they exist
            const safeSelect = 'ID,Title,Approval_Status,Status,Approval_Level,Current_Approver,Approval_History,Approval_Started_On,Approval_Completed_On,Created,Modified,PO_Number,PR_Number,Project_Code,Vendor,TotalAmount,Amount';
            const safeExpand = 'Project_Code,Vendor';

            const url = `${base}?$select=${safeSelect}&$expand=${safeExpand}`;

            let response = await spHttpClient.get(url, SPHttpClient.configurations.v1);

            if (!response.ok) {
                // Fallback to the absolute minimal query (no custom fields at all) – prevents 400 when columns are missing
                const minimalUrl = `${base}`;
                response = await spHttpClient.get(minimalUrl, SPHttpClient.configurations.v1);
                if (!response.ok) {
                    throw new Error(`Failed to fetch record: ${response.statusText}`);
                }
            }

            const data = await response.json();
            setState((prev) => ({
                ...prev,
                record: data,
                isLoading: false,
            }));
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to fetch approval details';
            setState((prev) => ({
                ...prev,
                isLoading: false,
                error: errorMessage,
            }));
        }
    };

    const handleRefresh = (): void => {
        fetchRecord();
        if (onRefresh) {
            onRefresh();
        }
    };

    // Get the approval status color
    const getStatusColor = (status: string): string => {
        switch (status?.toLowerCase()) {
            case 'approved':
            case 'completed':
                return theme.palette.green;
            case 'rejected':
                return theme.palette.red;
            case 'pending':
                return theme.palette.orange;
            default:
                return theme.palette.neutralSecondary;
        }
    };

    // Get the approval status background color
    const getStatusBackgroundColor = (status: string): string => {
        switch (status?.toLowerCase()) {
            case 'approved':
            case 'completed':
                return '#dff6dd';
            case 'rejected':
                return '#fde7e9';
            case 'pending':
                return '#fff4ce';
            default:
                return theme.palette.neutralLighter;
        }
    };

    // Get module display name from list name
    const getModuleDisplayName = (listName: string): string => {
        const moduleMap: { [key: string]: string } = {
            'PRC_Material_Request_Register': 'Material Request',
            'PRC_Purchase_Requisition_Register': 'Purchase Requisition',
            'PRC_Purchase_Order_Register': 'Purchase Order',
            'FIN_Payment_Request_Register': 'Payment Request',
            'FIN_Expense_Register': 'Expense',
            'ENT_Project_Master': 'Project',
        };
        return moduleMap[listName] || listName;
    };

    // Parse approval history from the record
    const getApprovalHistory = (): Array<{ action: string; comment: string; timestamp?: string; date?: string; approver: string }> => {
        if (!state.record?.Approval_History) return [];
        try {
            const history = JSON.parse(state.record.Approval_History);
            return Array.isArray(history) ? history : [];
        } catch {
            return [];
        }
    };

    // Determine approval stage based on history + record fields
    // PO uses amount-driven Approval_Level (can start at 1 or 2/CFO).
    // MR/PR use incremental approval.
    // GRN is 1-stage + "Notify Finance for Payment" as the follow-up action.
    const getApprovalStage = (): { stage: number; label: string } => {
        const history = getApprovalHistory();
        const approvalStatus = (state.record?.Approval_Status || state.record?.Status || '').toLowerCase();
        const approvalLevel = state.record?.Approval_Level || 0;
        const hasCurrentApprover = !!state.record?.Current_Approver;
        const isGRN = listName.includes('GRN') || listName.includes('Goods_Received') || listName.toLowerCase().includes('grn');

        if (approvalStatus === 'rejected') {
            return { stage: 0, label: 'Rejected' };
        }

        if (isGRN) {
            // GRN specific flow: Stage 1 = GRN Approver, then "Notify Finance for Payment"
            if (approvalStatus === 'approved' || approvalStatus === 'completed' || !hasCurrentApprover) {
                return { stage: 2, label: 'Notify Finance for Payment' };
            }
            return { stage: 1, label: 'Stage 1 (GRN Approver)' };
        }

        // Completed only when explicitly approved or no pending approver
        if (approvalStatus === 'approved' || approvalStatus === 'completed' || !hasCurrentApprover) {
            return { stage: 3, label: 'Completed' };
        }

        // For PO, trust Approval_Level directly (it can be 1 or 2 from creation)
        if (approvalLevel === 2) {
            return { stage: 2, label: 'Stage 2 (CFO)' };
        }
        if (approvalLevel === 1) {
            return { stage: 1, label: 'Stage 1' };
        }

        if (history.length === 0) {
            return { stage: 1, label: 'Stage 1' };
        }

        const hasRejected = history.some((h) => (h.action || '').toLowerCase() === 'rejected' || (h.action || '').toLowerCase() === 'reject');
        if (hasRejected) {
            return { stage: 0, label: 'Rejected' };
        }

        const approvedCount = history.filter((h) => (h.action || '').toLowerCase() === 'approved' || (h.action || '').toLowerCase() === 'approve').length;
        if (approvedCount >= 1 || approvalLevel > 1) {
            return { stage: 2, label: 'Stage 2' };
        }
        return { stage: 1, label: 'Stage 1' };
    };

    // Get last rejection comment
    const getRejectionComment = (): string | undefined => {
        const history = getApprovalHistory();
        const rejections = history.filter((h) => h.action === 'reject');
        if (rejections.length > 0) {
            return rejections[rejections.length - 1].comment;
        }
        return undefined;
    };

    // Format date for display
    const formatDate = (dateString: string | undefined): string => {
        if (!dateString) return 'N/A';
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'short',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
            });
        } catch {
            return 'N/A';
        }
    };

    const classNames = mergeStyleSets({
        panelContent: {
            padding: '20px',
        },
        headerSection: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '20px',
        },
        headerTitle: {
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
        },
        refreshButton: {
            marginBottom: '16px',
        },
        section: {
            marginBottom: '24px',
        },
        sectionTitle: {
            fontWeight: 600,
            marginBottom: '16px',
            color: theme.palette.neutralPrimary,
            fontSize: '16px',
        },
        fieldRow: {
            display: 'flex',
            flexWrap: 'wrap',
            gap: '16px',
            marginBottom: '12px',
        },
        fieldColumn: {
            flex: '1 1 200px',
            minWidth: '150px',
        },
        fieldLabel: {
            fontWeight: 600,
            marginBottom: '4px',
            color: theme.palette.neutralPrimary,
            fontSize: '12px',
            display: 'block',
        },
        fieldValue: {
            color: theme.palette.neutralSecondary,
            fontSize: '14px',
            wordBreak: 'break-word',
        },
        statusBadge: {
            display: 'inline-block',
            padding: '4px 12px',
            borderRadius: '4px',
            fontSize: '12px',
            fontWeight: 600,
            textTransform: 'uppercase',
        },
        stageContainer: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'flex-start',
            gap: '8px',
            marginTop: '16px',
            flexWrap: 'wrap',
        },
        stageItem: {
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
        },
        stageCircle: {
            width: '32px',
            height: '32px',
            borderRadius: '50%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 600,
            fontSize: '14px',
        },
        stageLine: {
            width: '40px',
            height: '2px',
            backgroundColor: theme.palette.neutralLight,
        },
        stageLabel: {
            fontSize: '14px',
            fontWeight: 500,
        },
        pendingMessage: {
            padding: '12px 16px',
            backgroundColor: '#fff4ce',
            borderRadius: '4px',
            borderLeft: `4px solid ${theme.palette.orange}`,
            marginTop: '16px',
            color: '#ca5010',
            fontSize: '14px',
        },
        rejectionMessage: {
            padding: '12px 16px',
            backgroundColor: '#fde7e9',
            borderRadius: '4px',
            borderLeft: `4px solid ${theme.palette.red}`,
            marginTop: '16px',
            color: '#a80000',
            fontSize: '14px',
        },
        historyContainer: {
            backgroundColor: theme.palette.neutralLighterAlt,
            padding: '12px',
            borderRadius: '4px',
            maxHeight: '200px',
            overflowY: 'auto',
            border: `1px solid ${theme.palette.neutralLight}`,
        },
        historyItem: {
            padding: '8px',
            borderBottom: `1px solid ${theme.palette.neutralLight}`,
            fontSize: '12px',
            marginBottom: '8px',
            '&:last-child': {
                borderBottom: 'none',
            },
        },
        historyTimestamp: {
            color: theme.palette.neutralTertiary,
            fontSize: '11px',
            marginBottom: '4px',
        },
        historyComment: {
            color: theme.palette.neutralSecondary,
            fontStyle: 'italic',
            marginTop: '4px',
        },
        loadingContainer: {
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            padding: '40px',
            textAlign: 'center',
        },
        errorContainer: {
            padding: '16px',
            backgroundColor: '#fde7e9',
            borderRadius: '4px',
            borderLeft: `4px solid ${theme.palette.red}`,
            color: '#a80000',
        },
        recordInfo: {
            padding: '12px',
            backgroundColor: theme.palette.neutralLighter,
            borderRadius: '4px',
            marginBottom: '16px',
        },
    });

    // Determine if status is pending
    const isPending = state.record?.Approval_Status?.toLowerCase() === 'pending';
    const isRejected = state.record?.Approval_Status?.toLowerCase() === 'rejected';
    const stageInfo = getApprovalStage();

    const renderStageIndicator = () => {
        const stages = [
            { stage: 1, label: 'Stage 1' },
            { stage: 2, label: 'Stage 2' },
            { stage: 3, label: 'Completed' },
        ];

        return (
            <div className={classNames.stageContainer}>
                {stages.map((stage, index) => {
                    const isActive = stageInfo.stage >= stage.stage;
                    const isCurrent = stageInfo.stage === stage.stage;

                    return (
                        <React.Fragment key={stage.stage}>
                            <div className={classNames.stageItem}>
                                <div
                                    className={classNames.stageCircle}
                                    style={{
                                        backgroundColor: isActive
                                            ? theme.palette.themePrimary
                                            : theme.palette.neutralLight,
                                        color: isActive ? theme.palette.white : theme.palette.neutralSecondary,
                                        border: isCurrent ? `2px solid ${theme.palette.themeDark}` : 'none',
                                    }}
                                >
                                    {stage.stage === 3 ? '✓' : stage.stage}
                                </div>
                                <span
                                    className={classNames.stageLabel}
                                    style={{
                                        color: isActive
                                            ? theme.palette.neutralPrimary
                                            : theme.palette.neutralTertiary,
                                        fontWeight: isCurrent ? 600 : 400,
                                    }}
                                >
                                    {stage.label}
                                </span>
                            </div>
                            {index < stages.length - 1 && (
                                <div
                                    className={classNames.stageLine}
                                    style={{
                                        backgroundColor:
                                            stageInfo.stage > stage.stage
                                                ? theme.palette.themePrimary
                                                : theme.palette.neutralLight,
                                    }}
                                />
                            )}
                        </React.Fragment>
                    );
                })}
            </div>
        );
    };

    // Render loading state
    if (state.isLoading) {
        return (
            <Panel
                isOpen={isOpen}
                onDismiss={onDismiss}
                type={PanelType.medium}
                headerText="Track Approval"
                closeButtonAriaLabel="Close"
            >
                <div className={classNames.loadingContainer}>
                    <Spinner size={SpinnerSize.large} label="Loading approval details..." />
                </div>
            </Panel>
        );
    }

    // Render error state
    if (state.error) {
        return (
            <Panel
                isOpen={isOpen}
                onDismiss={onDismiss}
                type={PanelType.medium}
                headerText="Track Approval"
                closeButtonAriaLabel="Close"
            >
                <div className={classNames.errorContainer}>
                    <Text variant="medium">{state.error}</Text>
                </div>
            </Panel>
        );
    }

    // Render no record found
    if (!state.record) {
        return (
            <Panel
                isOpen={isOpen}
                onDismiss={onDismiss}
                type={PanelType.medium}
                headerText="Track Approval"
                closeButtonAriaLabel="Close"
            >
                <div className={classNames.errorContainer}>
                    <Text variant="medium">Record not found</Text>
                </div>
            </Panel>
        );
    }

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.medium}
            headerText={`Track Approval - ${getModuleDisplayName(listName)}`}
            closeButtonAriaLabel="Close"
            isLightDismiss
        >
            <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                <div className={classNames.panelContent}>
                    {/* Refresh Button */}
                    <div className={classNames.refreshButton}>
                        <IconButton
                            iconProps={{ iconName: 'Refresh' }}
                            onClick={handleRefresh}
                            title="Refresh approval status"
                            ariaLabel="Refresh approval status"
                        />
                    </div>

                    {/* Record Info Section */}
                    <div className={classNames.section}>
                        <div className={classNames.sectionTitle}>Record Information</div>
                        <div className={classNames.recordInfo}>
                            <div className={classNames.fieldRow}>
                                <div className={classNames.fieldColumn}>
                                    <Label className={classNames.fieldLabel}>Record ID</Label>
                                    <Text className={classNames.fieldValue}>#{state.record.ID}</Text>
                                </div>
                                <div className={classNames.fieldColumn}>
                                    <Label className={classNames.fieldLabel}>Reference</Label>
                                    <Text className={classNames.fieldValue}>
                                        {state.record.Title || state.record.Payment_Number || state.record.PR_Number || state.record.PO_Number || state.record.MR_Number || 'N/A'}
                                    </Text>
                                </div>
                            </div>
                        </div>
                    </div>

                    {/* Approval Status Section */}
                    <div className={classNames.section}>
                        <div className={classNames.sectionTitle}>Approval Status</div>
                        <div className={classNames.fieldRow}>
                            <div className={classNames.fieldColumn}>
                                <Label className={classNames.fieldLabel}>Status</Label>
                                <span
                                    className={classNames.statusBadge}
                                    style={{
                                        backgroundColor: getStatusBackgroundColor(state.record.Approval_Status),
                                        color: getStatusColor(state.record.Approval_Status),
                                    }}
                                >
                                    {state.record.Approval_Status || 'Pending'}
                                </span>
                            </div>
                            <div className={classNames.fieldColumn}>
                                <Label className={classNames.fieldLabel}>Current Approver</Label>
                                <Text className={classNames.fieldValue}>
                                    {state.record.Current_Approver || state.record.Current_ApproverId?.Title || 'Not Assigned'}
                                </Text>
                            </div>
                        </div>

                        {/* Stage Indicator */}
                        {renderStageIndicator()}

                        {/* Pending Message */}
                        {isPending && state.record.Current_Approver && (
                            <div className={classNames.pendingMessage}>
                                <Text variant="medium">
                                    Waiting for approval from {state.record.Current_Approver}
                                </Text>
                            </div>
                        )}

                        {/* Rejection Comments */}
                        {isRejected && getRejectionComment() && (
                            <div className={classNames.rejectionMessage}>
                                <Text variant="medium" style={{ fontWeight: 600 }}>
                                    Rejection Reason:
                                </Text>
                                <Text variant="medium" style={{ display: 'block', marginTop: '4px' }}>
                                    {getRejectionComment()}
                                </Text>
                            </div>
                        )}
                    </div>

                    {/* Comments Section */}
                    {state.record.Comments && (
                        <div className={classNames.section}>
                            <div className={classNames.sectionTitle}>Approval Comments</div>
                            <div
                                style={{
                                    padding: '12px',
                                    backgroundColor: theme.palette.neutralLighter,
                                    borderRadius: '4px',
                                }}
                            >
                                <Text className={classNames.fieldValue}>{state.record.Comments}</Text>
                            </div>
                        </div>
                    )}

                    {/* Dates Section */}
                    <div className={classNames.section}>
                        <div className={classNames.sectionTitle}>Timeline</div>
                        <div className={classNames.fieldRow}>
                            <div className={classNames.fieldColumn}>
                                <Label className={classNames.fieldLabel}>Created Date</Label>
                                <Text className={classNames.fieldValue}>
                                    {formatDate(state.record.Created)}
                                </Text>
                            </div>
                            <div className={classNames.fieldColumn}>
                                <Label className={classNames.fieldLabel}>Last Modified</Label>
                                <Text className={classNames.fieldValue}>
                                    {formatDate(state.record.Modified)}
                                </Text>
                            </div>
                        </div>
                    </div>

                    {/* Approval History Section */}
                    {getApprovalHistory().length > 0 && (
                        <div className={classNames.section}>
                            <div className={classNames.sectionTitle}>Approval History</div>
                            <div className={classNames.historyContainer}>
                                {getApprovalHistory().map((action, index) => (
                                    <div key={index} style={{ marginBottom: '12px', paddingBottom: '12px', borderBottom: `1px solid ${theme.palette.neutralLight}` }}>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
                                            <span
                                                style={{
                                                    display: 'inline-block',
                                                    padding: '2px 8px',
                                                    borderRadius: '3px',
                                                    fontSize: '11px',
                                                    fontWeight: 600,
                                                    textTransform: 'uppercase',
                                                    backgroundColor: action.action === 'approve' ? theme.palette.green : theme.palette.red,
                                                    color: theme.palette.white,
                                                }}
                                            >
                                                {action.action}
                                            </span>
                                            <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                                                by {action.approver}
                                            </span>
                                        </div>
                                         <div style={{ fontSize: '11px', color: theme.palette.neutralTertiary }}>
                                             {formatDate(action.date || action.timestamp)}
                                         </div>
                                        {action.comment && (
                                            <div style={{ fontSize: '12px', fontStyle: 'italic', color: theme.palette.neutralSecondary, marginTop: '4px' }}>
                                                "{action.comment}"
                                            </div>
                                        )}
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}
                </div>
            </ScrollablePane>
        </Panel>
    );
};

export default ApprovalTrackerPanel;