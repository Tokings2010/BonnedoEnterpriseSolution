import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
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

// Approved Payments Component
interface IFinanceApprovedPaymentsProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const FinanceApprovedPayments: React.FC<IFinanceApprovedPaymentsProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [selectedRecord, setSelectedRecord] = React.useState<IListItem | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [refreshKey, setRefreshKey] = React.useState(0);
    const [paymentDate, setPaymentDate] = React.useState<string>(new Date().toISOString().split('T')[0]);
    const [paymentMethod, setPaymentMethod] = React.useState<string>('Bank Transfer');
    const [transactionRef, setTransactionRef] = React.useState('');
    const [saving, setSaving] = React.useState(false);
    const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);

    const approvedPaymentColumns: IDataGridColumn[] = [
        { key: 'Payment_Number', name: 'Payment Number', fieldName: 'Payment_Number', minWidth: 120, isResizable: true },
        { key: 'Vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150, isResizable: true },
        { key: 'Project_Code', name: 'Project Code', fieldName: 'Project_Code', minWidth: 100, isResizable: true },
        { key: 'Amount', name: 'Amount', fieldName: 'Amount', minWidth: 120, isResizable: true },
        { key: 'Payment_Status', name: 'Payment Status', fieldName: 'Payment_Status', minWidth: 100, isResizable: true },
        { key: 'Payment_Date', name: 'Payment Date', fieldName: 'Payment_Date', minWidth: 120, isResizable: true },
        { key: 'Processed_By', name: 'Processed By', fieldName: 'Processed_By', minWidth: 120, isResizable: true },
    ];

    const filterQuery = "Approval_Status eq 'Approved'";

    const handleRowDoubleClick = (record: IListItem): void => {
        setSelectedRecord(record);
        setIsPanelOpen(true);
        // Pre-fill with existing data if available
        if (record.Payment_Date) {
            setPaymentDate(record.Payment_Date.split('T')[0]);
        }
        if (record.Payment_Method) {
            setPaymentMethod(record.Payment_Method);
        }
        if (record.Transaction_Reference) {
            setTransactionRef(record.Transaction_Reference);
        }
    };

    const handleRefresh = (): void => {
        setRefreshKey(prev => prev + 1);
        if (onRefresh) onRefresh();
    };

    const handleMarkAsPaid = async (): Promise<void> => {
        if (!selectedRecord) return;
        setSaving(true);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/lists/getByTitle('FIN_Payment_Request_Register')/items(${selectedRecord.ID})`;

            await spHttpClient.post(url, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*',
                },
                body: JSON.stringify({
                    Payment_Status: 'Paid',
                    Payment_Date: paymentDate,
                    Payment_Method: paymentMethod,
                    Transaction_Reference: transactionRef,
                    Processed_By: pageContext.user.displayName,
                }),
            });

            setMessage({ type: MessageBarType.success, text: 'Payment marked as paid successfully!' });
            setIsPanelOpen(false);
            handleRefresh();
        } catch (error) {
            setMessage({ type: MessageBarType.error, text: 'Failed to update payment status' });
        }
        setSaving(false);
    };

    const classNames = mergeStyleSets({
        container: { padding: '20px', height: '100%' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' },
        headerTitle: { flex: 1, minWidth: '200px' },
        headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
        panelContent: { padding: '20px' },
        detailRow: { display: 'flex', marginBottom: '12px' },
        detailLabel: { fontWeight: 600, width: '140px', color: theme.palette.neutralSecondary },
        detailValue: { flex: 1 },
        formGroup: { marginBottom: '16px' },
        buttonGroup: { display: 'flex', gap: '8px', marginTop: '16px' },
    });

    return (
        <div className={classNames.container}>
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Approved Payments
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        View approved payment transactions
                    </Text>
                </div>
                <div className={classNames.headerActions}>
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
                    columns={approvedPaymentColumns}
                    filterQuery={filterQuery}
                    pageSize={10}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onRowDoubleClick={handleRowDoubleClick}
                    showExport
                />
            </div>

            {/* Payment Processing Panel */}
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
                            <Text variant="large" style={{ fontWeight: 600 }}>Payment Processing</Text>
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
                            <span className={classNames.detailLabel}>Vendor:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Vendor}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Project Code:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Project_Code}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Amount:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Amount}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Current Status:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Payment_Status}</span>
                        </div>

                        <hr style={{ margin: '20px 0', border: 'none', borderTop: `1px solid ${theme.palette.neutralLight}` }} />

                        <Text variant="large" style={{ fontWeight: 600, marginBottom: '16px', display: 'block' }}>
                            Payment Details
                        </Text>

                        <div className={classNames.formGroup}>
                            <Label>Payment Date</Label>
                            <TextField
                                type="date"
                                value={paymentDate}
                                onChange={(_, value) => setPaymentDate(value || '')}
                            />
                        </div>

                        <div className={classNames.formGroup}>
                            <Label>Payment Method</Label>
                            <TextField
                                value={paymentMethod}
                                onChange={(_, value) => setPaymentMethod(value || '')}
                                placeholder="e.g., Bank Transfer, Cash, Cheque"
                            />
                        </div>

                        <div className={classNames.formGroup}>
                            <Label>Transaction Reference</Label>
                            <TextField
                                value={transactionRef}
                                onChange={(_, value) => setTransactionRef(value || '')}
                                placeholder="e.g., TRF-2024-001234"
                            />
                        </div>

                        <div className={classNames.buttonGroup}>
                            <PrimaryButton
                                text="Mark as Paid"
                                iconProps={{ iconName: 'Money' }}
                                onClick={handleMarkAsPaid}
                                disabled={saving || selectedRecord.Payment_Status === 'Paid'}
                            />
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

export default FinanceApprovedPayments;
