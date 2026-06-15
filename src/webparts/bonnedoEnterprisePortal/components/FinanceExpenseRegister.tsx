import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    DefaultButton,
    PrimaryButton,
    TextField,
    MessageBar,
    MessageBarType,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import DataGrid from './DataGrid';
import { IDataGridColumn } from './DataGrid';
import { IListItem } from '../services/SharePointService';
import ApprovalTrackerPanel from './ApprovalTrackerPanel';

// Expense Register Component
interface IFinanceExpenseRegisterProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    onNewExpense?: () => void;
    searchText?: string;
    isMobileView?: boolean;
}

const FinanceExpenseRegister: React.FC<IFinanceExpenseRegisterProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    onNewExpense,
    searchText = '',
    isMobileView = false
}) => {
    const theme = getTheme();
    const [selectedRecord, setSelectedRecord] = React.useState<IListItem | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [refreshKey, setRefreshKey] = React.useState(0);
    const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);
    const [isApprovalTrackerOpen, setIsApprovalTrackerOpen] = React.useState(false);

    const expenseColumns: IDataGridColumn[] = [
        { key: 'Expense_Number', name: 'Expense Number', fieldName: 'Expense_Number', minWidth: 120, isResizable: true },
        { key: 'Project_Code', name: 'Project Code', fieldName: 'Project_Code', minWidth: 100, isResizable: true },
        { key: 'Expense_Type', name: 'Expense Type', fieldName: 'Expense_Type', minWidth: 120, isResizable: true },
        { key: 'Amount', name: 'Amount', fieldName: 'Amount', minWidth: 120, isResizable: true },
        { key: 'Recorded_By', name: 'Recorded By', fieldName: 'Recorded_By', minWidth: 120, isResizable: true },
        { key: 'Approval_Status', name: 'Status', fieldName: 'Approval_Status', minWidth: 100, isResizable: true },
        { key: 'Created', name: 'Created', fieldName: 'Created', minWidth: 120, isResizable: true },
    ];

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

    const filterQuery = undefined;

    const classNames = mergeStyleSets({
        container: { padding: '20px', height: '100%' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' },
        headerTitle: { flex: 1, minWidth: '200px' },
        headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },

        panelContent: { padding: '20px' },
        detailRow: { display: 'flex', marginBottom: '12px' },
        detailLabel: { fontWeight: 600, width: '140px', color: theme.palette.neutralSecondary },
        detailValue: { flex: 1 },
    });

    // Get status badge styles
    const getStatusBadge = (status: string) => {
        let bgColor = '#f3f2f1';
        let textColor = '#605e5c';

        switch (status?.toLowerCase()) {
            case 'approved':
                bgColor = '#dff6dd';
                textColor = '#107c10';
                break;
            case 'rejected':
                bgColor = '#fde7e9';
                textColor = '#a80000';
                break;
            case 'pending':
                bgColor = '#fff4ce';
                textColor = '#ca5010';
                break;
        }

        return {
            display: 'inline-block',
            padding: '2px 8px',
            borderRadius: '10px',
            fontSize: '11px',
            fontWeight: 600,
            textTransform: 'uppercase',
            backgroundColor: bgColor,
            color: textColor,
        };
    };

    return (
        <div className={classNames.container}>
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Expense Register
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Track and manage expenses
                    </Text>
                </div>
                <div className={classNames.headerActions}>
                    {onNewExpense && (
                        <PrimaryButton
                            text="+ New Expense"
                            onClick={onNewExpense}
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
                    listName="FIN_Expense_Register"
                    columns={expenseColumns}
                    filterQuery={filterQuery}
                    pageSize={10}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onRowDoubleClick={handleRowDoubleClick}
                />
            </div>

            {/* Expense Details Panel */}
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
                            <Text variant="large" style={{ fontWeight: 600 }}>Expense Details</Text>
                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                onClick={() => setIsPanelOpen(false)}
                            />
                        </div>
                    </div>

                    <div className={classNames.panelContent}>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Expense Number:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Expense_Number}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Project Code:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Project_Code}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Expense Type:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Expense_Type}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Amount:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Amount}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Recorded By:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Recorded_By}</span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Status:</span>
                            <span style={getStatusBadge(selectedRecord.Approval_Status)}>
                                {selectedRecord.Approval_Status}
                            </span>
                        </div>
                        <div className={classNames.detailRow}>
                            <span className={classNames.detailLabel}>Created:</span>
                            <span className={classNames.detailValue}>{selectedRecord.Created}</span>
                        </div>
                        {selectedRecord.Notes && (
                            <div className={classNames.detailRow} style={{ flexDirection: 'column' }}>
                                <span className={classNames.detailLabel}>Notes:</span>
                                <div style={{
                                    backgroundColor: theme.palette.neutralLighterAlt,
                                    padding: '12px',
                                    borderRadius: '4px',
                                    marginTop: '8px'
                                }}>
                                    {selectedRecord.Notes}
                                </div>
                            </div>
                        )}

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
                    listName="FIN_Expense_Register"
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onDismiss={closeApprovalTracker}
                    onRefresh={handleApprovalRefresh}
                />
            )}
        </div>
    );
};

export default FinanceExpenseRegister;
