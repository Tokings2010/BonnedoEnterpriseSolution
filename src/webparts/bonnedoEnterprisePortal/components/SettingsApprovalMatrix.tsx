import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    PrimaryButton,
    MessageBar,
    MessageBarType,
} from '@fluentui/react';
import { IDropdownOption } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import DataGrid from './DataGrid';
import { IDataGridColumn } from './DataGrid';
import { IListItem } from '../services/SharePointService';
import ApprovalMatrixForm from './ApprovalMatrixForm';
import ApprovalMatrixDetailsPanel from './ApprovalMatrixDetailsPanel';

// Helper function to extract value from either string or lookup object
const extractCellValue = (item: IListItem, fieldName: string): string => {
    const value = item[fieldName];
    if (!value) return '';
    // If it's an object (from expandQuery), try to get the Title
    if (typeof value === 'object' && value !== null) {
        return (value as { Title?: string }).Title || (value as { title?: string }).title || '';
    }
    // Otherwise return as string
    return String(value);
};

// Settings Approval Matrix Component
interface ISettingsApprovalMatrixProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
    onRefresh?: () => void;
    onNewApproval?: () => void;
    isMobileView?: boolean;
}

const SettingsApprovalMatrix: React.FC<ISettingsApprovalMatrixProps> = ({
    spHttpClient,
    pageContext,
    webPartContext,
    onRefresh,
    onNewApproval,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [selectedRecord, setSelectedRecord] = React.useState<IListItem | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [isEditMode, setIsEditMode] = React.useState(false);
    const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
    const [refreshKey, setRefreshKey] = React.useState(0);
    const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);
    const [moduleOptions, setModuleOptions] = React.useState<IDropdownOption[]>([]);
    const [approverRoleOptions, setApproverRoleOptions] = React.useState<IDropdownOption[]>([]);
    const [isLoadingMetadata, setIsLoadingMetadata] = React.useState(false);
    const [metadataError, setMetadataError] = React.useState<string | null>(null);

    const fetchChoiceColumns = React.useCallback(async () => {
        setIsLoadingMetadata(true);
        setMetadataError(null);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            const fieldsUrl = `${webUrl}/_api/web/lists/getByTitle('SYS_Approval_Matrix')/fields?$select=Title,TypeAsString,Choices&$filter=TypeAsString eq 'Choice'`;
            const response = await spHttpClient.get(fieldsUrl, SPHttpClient.configurations.v1);
            if (!response.ok) throw new Error('Failed to fetch field metadata');
            const data = await response.json();
            const fields = data.value || [];
            fields.forEach((field: any) => {
                const choices: string[] = field.Choices || [];
                const options: IDropdownOption[] = choices.map((c: string) => ({ key: c, text: c }));
                if (field.Title === 'Module') {
                    setModuleOptions(options);
                } else if (field.Title === 'Approver_Role') {
                    setApproverRoleOptions(options);
                }
            });
        } catch (err) {
            setMetadataError(err instanceof Error ? err.message : 'Failed to load choices');
        } finally {
            setIsLoadingMetadata(false);
        }
    }, [spHttpClient, pageContext]);

    React.useEffect(() => {
        fetchChoiceColumns();
    }, [fetchChoiceColumns]);

    const approvalMatrixColumns: IDataGridColumn[] = [
        { key: 'Module', name: 'Module', fieldName: 'Module', minWidth: 150, isResizable: true },
        { key: 'Min_Amount', name: 'Min Amount', fieldName: 'Min_Amount', minWidth: 120, isResizable: true },
        { key: 'Max_Amount', name: 'Max Amount', fieldName: 'Max_Amount', minWidth: 120, isResizable: true },
        { key: 'Stage', name: 'Stage', fieldName: 'Stage', minWidth: 80, isResizable: true },
        { key: 'Approver_Role', name: 'Approver Role', fieldName: 'Approver_Role', minWidth: 150, isResizable: true },
        {
            key: 'Approver_User',
            name: 'Approver User',
            fieldName: 'Approver_User',
            minWidth: 150,
            isResizable: true,
            onRender: (item?: IListItem) => {
                const value = item ? extractCellValue(item, 'Approver_User') : '';
                return <Text>{value || '-'}</Text>;
            }
        },
    ];

    const handleRowDoubleClick = (record: IListItem): void => {
        setSelectedRecord(record);
        setIsDetailsPanelOpen(true);
    };

    const handleDetailsPanelDismiss = (): void => {
        setIsDetailsPanelOpen(false);
        setSelectedRecord(undefined);
    };

    const handleDetailsPanelRefresh = (): void => {
        handleRefresh();
    };

    const handleNewClick = (): void => {
        setSelectedRecord(undefined);
        setIsEditMode(false);
        setIsPanelOpen(true);
    };

    const handleRefresh = (): void => {
        setRefreshKey(prev => prev + 1);
        onRefresh?.();
    };

    const handleFormDismiss = (): void => {
        setIsPanelOpen(false);
        setSelectedRecord(undefined);
        setIsEditMode(false);
    };

    const handleFormSuccess = (): void => {
        setMessage({
            type: MessageBarType.success,
            text: isEditMode ? 'Approval rule updated successfully!' : 'Approval rule created successfully!'
        });
        handleRefresh();
        // Clear message after 3 seconds
        setTimeout(() => setMessage(null), 3000);
    };

    const classNames = mergeStyleSets({
        container: { padding: '20px' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' },
        panelContent: { padding: '20px' },
        formGroup: { marginBottom: '16px' },
        buttonGroup: { display: 'flex', gap: '8px', marginTop: '16px' },
    });

    return (
        <div className={classNames.container}>
            {/* Tab Header */}
            <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                <div style={{ flex: 1, minWidth: '200px' }}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Approval Matrix
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Configure approval workflows and rules
                    </Text>
                </div>
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    <PrimaryButton
                        text="+ New Approval Rule"
                        onClick={handleNewClick}
                        iconProps={{ iconName: 'Add' }}
                    />
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
            {metadataError && (
                <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => setMetadataError(null)}>
                    {metadataError} - Using fallback options if available.
                </MessageBar>
            )}

            <div style={{ height: 'calc(100vh - 280px)' }}>
                <DataGrid
                    key={refreshKey}
                    listName="SYS_Approval_Matrix"
                    columns={approvalMatrixColumns}
                    pageSize={10}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onRowDoubleClick={handleRowDoubleClick}
                    expandQuery="Approver_User"
                />
            </div>

            {/* Edit/Create Panel using ApprovalMatrixForm */}
            {isPanelOpen && (
                <ApprovalMatrixForm
                    isOpen={isPanelOpen}
                    onDismiss={handleFormDismiss}
                    onSubmitSuccess={handleFormSuccess}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    webPartContext={webPartContext}
                    editMode={isEditMode}
                    editRecord={selectedRecord ? {
                        ID: selectedRecord.ID,
                        Module: selectedRecord.Module || '',
                        Min_Amount: parseFloat(selectedRecord.Min_Amount || '0'),
                        Max_Amount: parseFloat(selectedRecord.Max_Amount || '0'),
                        Stage: selectedRecord.Stage ? parseInt(selectedRecord.Stage) : undefined,
                        Approver_Role: selectedRecord.Approver_Role || '',
                        Approver_UserId: selectedRecord.Approver_UserId || 0,
                        Approver_User: selectedRecord.Approver_User
                    } : undefined}
                    moduleOptions={moduleOptions}
                    approverRoleOptions={approverRoleOptions}
                />
            )}

            {/* Details Panel using ApprovalMatrixDetailsPanel */}
            {isDetailsPanelOpen && selectedRecord && (
                <ApprovalMatrixDetailsPanel
                    isOpen={isDetailsPanelOpen}
                    approvalMatrix={{
                        ID: selectedRecord.ID,
                        Module: selectedRecord.Module || '',
                        Min_Amount: parseFloat(selectedRecord.Min_Amount || '0'),
                        Max_Amount: parseFloat(selectedRecord.Max_Amount || '0'),
                        Stage: selectedRecord.Stage ? parseInt(selectedRecord.Stage) : undefined,
                        Approver_Role: selectedRecord.Approver_Role || '',
                        Approver_UserId: selectedRecord.Approver_UserId || 0,
                        Approver_User: selectedRecord.Approver_User
                    }}
                    onDismiss={handleDetailsPanelDismiss}
                    onRefresh={handleDetailsPanelRefresh}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    webPartContext={webPartContext}
                />
            )}
        </div>
    );
};

export default SettingsApprovalMatrix;
