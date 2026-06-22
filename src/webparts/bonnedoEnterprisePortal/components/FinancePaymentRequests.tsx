import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    DefaultButton,
    PrimaryButton,
    TextField,
    Label,
    MessageBar,
    MessageBarType,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import DataGrid from './EnhancedDataGrid';
import { IDataGridColumn } from './EnhancedDataGrid';
import { IListItem } from '../services/SharePointService';
import ApprovalTrackerPanel from './ApprovalTrackerPanel';

// Payment Requests Component
interface IFinancePaymentRequestsProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    onNewRequest?: () => void;
    searchText?: string;
    isMobileView?: boolean;
}

const FinancePaymentRequests: React.FC<IFinancePaymentRequestsProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    onNewRequest,
    searchText = '',
    isMobileView = false
}) => {
    const theme = getTheme();
    const [selectedRecord, setSelectedRecord] = React.useState<IListItem | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [refreshKey, setRefreshKey] = React.useState(0);
    const currentStatusFilter = 'All';
    const [approvalComment, setApprovalComment] = React.useState('');
    const [saving, setSaving] = React.useState(false);
    const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);
    const [isApprovalTrackerOpen, setIsApprovalTrackerOpen] = React.useState(false);

    const paymentColumns: IDataGridColumn[] = [
        { key: 'Payment_Number', name: 'Payment Number', fieldName: 'Payment_Number', minWidth: 120, isResizable: true },
        { key: 'Project_Code', name: 'Project Code', fieldName: 'Project_Code', minWidth: 100, isResizable: true },
        { key: 'Vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150, isResizable: true },
        { key: 'PR_Number', name: 'PR Number', fieldName: 'PR_Number', minWidth: 100, isResizable: true },
        { key: 'PO_Number', name: 'PO Number', fieldName: 'PO_Number', minWidth: 100, isResizable: true },
        { key: 'Amount', name: 'Amount', fieldName: 'Amount', minWidth: 120, isResizable: true },
        { key: 'Requested_By', name: 'Requested By', fieldName: 'Requested_By', minWidth: 120, isResizable: true },
        { key: 'Current_Approver', name: 'Current Approver', fieldName: 'Current_Approver', minWidth: 120, isResizable: true },
        { key: 'Approval_Status', name: 'Status', fieldName: 'Approval_Status', minWidth: 100, isResizable: true },
        { key: 'Payment_Status', name: 'Payment Status', fieldName: 'Payment_Status', minWidth: 100, isResizable: true },
        { key: 'Created', name: 'Created', fieldName: 'Created', minWidth: 120, isResizable: true },
    ];

    // Build filter query combining status and search
    const statusFilter = currentStatusFilter === 'All' ? '' : `Approval_Status eq '${currentStatusFilter}'`;

    const filterParts: string[] = [];
    if (statusFilter) filterParts.push(statusFilter);

    const filterQuery = filterParts.length > 0 ? filterParts.join(' and ') : undefined;

    const handleRowDoubleClick = (record: IListItem): void => {
        setSelectedRecord(record);
        setIsPanelOpen(true);
    };

    const handleRefresh = (): void => {
        setRefreshKey(prev => prev + 1);
        if (onRefresh) onRefresh();
    };

    const openApprovalTracker = (): void => {
        setIsApprovalTrackerOpen(true);
    };

    const closeApprovalTracker = (): void => {
        setIsApprovalTrackerOpen(false);
    };

    const handleApprovalRefresh = (): void => {
        handleRefresh();
    };

    const handleApprove = async (): Promise<void> => {
        if (!selectedRecord) return;
        setSaving(true);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/lists/getByTitle('FIN_Payment_Request_Register')/items(${selectedRecord.ID})`;

            const history = selectedRecord.Approval_History || '';
            const newHistory = history + `\n${new Date().toISOString()} - Approved by ${pageContext.user.displayName}: ${approvalComment}`;

            await spHttpClient.post(url, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*',
                },
                body: JSON.stringify({
                    Approval_Status: 'Approved',
                    Approval_History: newHistory,
                }),
            });

            setMessage({ type: MessageBarType.success, text: 'Payment request approved successfully!' });
            setIsPanelOpen(false);
            handleRefresh();
        } catch (error) {
            setMessage({ type: MessageBarType.error, text: 'Failed to approve payment request' });
        }
        setSaving(false);
    };

    const handleReject = async (): Promise<void> => {
        if (!selectedRecord) return;
        setSaving(true);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/lists/getByTitle('FIN_Payment_Request_Register')/items(${selectedRecord.ID})`;

            const history = selectedRecord.Approval_History || '';
            const newHistory = history + `\n${new Date().toISOString()} - Rejected by ${pageContext.user.displayName}: ${approvalComment}`;

            await spHttpClient.post(url, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*',
                },
                body: JSON.stringify({
                    Approval_Status: 'Rejected',
                    Approval_History: newHistory,
                }),
            });

            setMessage({ type: MessageBarType.success, text: 'Payment request rejected!' });
            setIsPanelOpen(false);
            handleRefresh();
        } catch (error) {
            setMessage({ type: MessageBarType.error, text: 'Failed to reject payment request' });
        }
        setSaving(false);
    };

    const handleReturn = async (): Promise<void> => {
        if (!selectedRecord) return;
        setSaving(true);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/lists/getByTitle('FIN_Payment_Request_Register')/items(${selectedRecord.ID})`;

            const history = selectedRecord.Approval_History || '';
            const newHistory = history + `\n${new Date().toISOString()} - Returned by ${pageContext.user.displayName}: ${approvalComment}`;

            await spHttpClient.post(url, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*',
                },
                body: JSON.stringify({
                    Approval_Status: 'Returned',
                    Approval_History: newHistory,
                }),
            });

            setMessage({ type: MessageBarType.success, text: 'Payment request returned!' });
            setIsPanelOpen(false);
            handleRefresh();
        } catch (error) {
            setMessage({ type: MessageBarType.error, text: 'Failed to return payment request' });
        }
        setSaving(false);
    };

    const classNames = mergeStyleSets({
        container: { padding: '20px', height: '100%' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' },
        headerTitle: { flex: 1, minWidth: '200px' },
        headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
        filterGroup: { display: 'flex', gap: '8px', alignItems: 'center' },

        panelContent: { padding: '20px' },
        detailRow: { display: 'flex', marginBottom: '12px' },
        detailLabel: { fontWeight: 600, width: '140px', color: theme.palette.neutralSecondary },
        detailValue: { flex: 1 },
        buttonGroup: { display: 'flex', gap: '8px', marginTop: '16px' },
        commentField: { marginTop: '16px' },
    });

    return (
        <div className={classNames.container}>
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Payment Requests
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Manage and track payment requests
                    </Text>
                </div>
                <div className={classNames.headerActions}>
                    {onNewRequest && (
                        <PrimaryButton
                            text="+ New Payment Request"
                            onClick={onNewRequest}
                            iconProps={{ iconName: 'Add' }}
                        />
                    )}
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={handleRefresh}
                        title="Refresh data"
                        ariaLabel="Refresh data"
                    />
                </div>
            </div>

            {message && (
                <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
                    {message.text}
                </MessageBar>
            )}

            <div style={{ height: 'calc(100vh - 220px)' }}>
                <DataGrid
                    key={refreshKey}
                    listName="FIN_Payment_Request_Register"
                    columns={paymentColumns}
                    filterQuery={filterQuery}
                    pageSize={10}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onRowDoubleClick={handleRowDoubleClick}
                    showExport
                />
            </div>

            {/* Payment Request Details Panel */}
            {isPanelOpen && selectedRecord && (
                <div style={{
                    position: 'fixed',
                    top: 0,
                    right: 0,
                    width: '500px',
                    height: '100vh',
                    backgroundColor: theme.palette.white,
                    boxShadow: theme.effects.elevation16,
                    zIndex: 1000,
                    overflowY: 'auto',
                    borderLeft: `1px solid ${theme.palette.neutralLight}`,
                }}>
                    <div style={{ padding: '20px', borderBottom: `1px solid ${theme.palette.neutralLight}` }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <Text variant="large" style={{ fontWeight: 600 }}>Payment Request Details</Text>
                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                onClick={() => setIsPanelOpen(false)}
                            />
                        </div>
                    </div>

                    <div className={classNames.panelContent}>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Payment Number:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Payment_Number}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Project Code:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Project_Code}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Vendor:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Vendor}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>PR Number:</span>
                            <span className={classNames.detailValue}>{selectedRecord.PR_Number}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>PO Number:</span>
                            <span className={classNames.detailValue}>{selectedRecord.PO_Number}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Amount:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Amount}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Requested By:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Requested_By}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Current Approver:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Current_Approver}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Status:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Approval_Status}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Payment Status:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Payment_Status}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Created:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Created}</span>
                        </div>
                        {selectedRecord.Approval_History && (
                            <div className={classNames.detailRow} style={{ flexDirection: 'column' }}>
                                <span className={classNames.detailLabel}>Approval History:</span>
                                <pre style={{
                                    backgroundColor: theme.palette.neutralLighterAlt,
                                    padding: '8px',
                                    borderRadius: '4px',
                                    whiteSpace: 'pre-wrap',
                                    fontSize: '12px'
                                }}>
                                    {selectedRecord.Approval_History}
                                </pre>
                            </div>
                        )}

                        <div className={classNames.commentField}>
                            <Label>Approval Comments</Label>
                            <TextField
                                multiline
                                rows={3}
                                value={approvalComment}
                                onChange={(_, value) => setApprovalComment(value || '')}
                                placeholder="Enter your comments..."
                            />
                        </div>

                        <div className={classNames.buttonGroup}>
                            <PrimaryButton
                                text="Approve"
                                iconProps={{ iconName: 'CheckMark' }}
                                onClick={handleApprove}
                                disabled={saving || selectedRecord.Approval_Status === 'Approved'}
                                style={{ backgroundColor: '#107c10' }}
                            />
                            <PrimaryButton
                                text="Reject"
                                iconProps={{ iconName: 'Cancel' }}
                                onClick={handleReject}
                                disabled={saving || selectedRecord.Approval_Status === 'Rejected'}
                                style={{ backgroundColor: '#a80000' }}
                            />
                            <DefaultButton
                                text="Return"
                                iconProps={{ iconName: 'Undo' }}
                                onClick={handleReturn}
                                disabled={saving}
                            />
                        </div>

                        {/* Track Approval Button */}
                        <div style={{ marginTop: '16px', paddingTop: '16px', borderTop: `1px solid ${theme.palette.neutralLight}` }}>
                            <DefaultButton
                                text="Track Approval"
                                iconProps={{ iconName: 'CheckCircle' }}
                                onClick={openApprovalTracker}
                                style={{ width: '100%' }}
                            />
                        </div>
                    </div>
                </div>
            )}

            {/* Approval Tracker Panel */}
            {selectedRecord && isApprovalTrackerOpen && (
                <ApprovalTrackerPanel
                    isOpen={isApprovalTrackerOpen}
                    recordId={selectedRecord.ID}
                    listName="FIN_Payment_Request_Register"
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onDismiss={closeApprovalTracker}
                    onRefresh={handleApprovalRefresh}
                />
            )}
        </div>
    );
};

export default FinancePaymentRequests;
