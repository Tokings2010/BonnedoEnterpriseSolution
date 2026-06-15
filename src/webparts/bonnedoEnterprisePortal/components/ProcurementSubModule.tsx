import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    PrimaryButton,
    IconButton,
    TextField,
    Label,
    DefaultButton,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import DataGrid from './DataGrid';
import { IDataGridColumn } from './DataGrid';
import MaterialRequestForm from './MaterialRequestModule';
import PurchaseRequisitionForm from './PurchaseRequisitionModule';
import PurchaseOrderForm from './PurchaseOrderModule';
import GoodsReceivedNoteForm from './GoodsReceivedNoteModule';
import ApprovalTrackerPanel from './ApprovalTrackerPanel';
import ProcurementDetailsPanel from './ProcurementDetailsPanel';

export interface IProcurementSubModuleProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    userDisplayName: string;
    recordType: 'MR' | 'PR' | 'PO' | 'GRN';
    webPartContext?: WebPartContext;
}

interface IProcurementRecord {
    ID: number;
    Title: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    [key: string]: any;
}

const ProcurementSubModule: React.FC<IProcurementSubModuleProps> = ({
    spHttpClient,
    pageContext,
    userDisplayName,
    recordType,
    webPartContext,
}) => {
    const theme = getTheme();
    const [selectedRecord, setSelectedRecord] = React.useState<IProcurementRecord | undefined>(undefined);
    const [isFormPanelOpen, setIsFormPanelOpen] = React.useState(false);
    const [refreshKey, setRefreshKey] = React.useState(0);
    const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
    const [isApprovalTrackerOpen, setIsApprovalTrackerOpen] = React.useState(false);

    const classNames = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            backgroundColor: theme.palette.white,
        },
        header: {
            marginBottom: '20px',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '12px',
        },
        headerTitle: {
            flex: 1,
            minWidth: '200px',
        },
        headerActions: {
            display: 'flex',
            gap: '8px',
            flexWrap: 'wrap',
        },

        gridContainer: {
            flex: 1,
            overflow: 'hidden',
            display: 'flex',
            flexDirection: 'column',
        },
        detailsPanel: {
            position: 'fixed',
            right: 0,
            top: 0,
            bottom: 0,
            width: '400px',
            backgroundColor: theme.palette.white,
            boxShadow: theme.effects.elevation16,
            padding: '20px',
            overflow: 'auto',
            zIndex: 1000,
            borderLeft: `1px solid ${theme.palette.neutralLight}`,
        },
        detailsHeader: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '20px',
            paddingBottom: '12px',
            borderBottom: `2px solid ${theme.palette.themePrimary}`,
        },
        detailsContent: {
            display: 'flex',
            flexDirection: 'column',
            gap: '16px',
        },
        detailsField: {
            display: 'flex',
            flexDirection: 'column',
            gap: '4px',
        },
        detailsLabel: {
            fontWeight: 600,
            color: theme.palette.neutralSecondary,
            fontSize: '12px',
            textTransform: 'uppercase',
        },
        detailsValue: {
            color: theme.palette.neutralPrimary,
            fontSize: '14px',
        },
    });

    const getListName = (): string => {
        switch (recordType) {
            case 'MR':
                return 'PRC_Material_Request_Register';
            case 'PR':
                return 'PRC_Purchase_Requisition_Register';
            case 'PO':
                return 'PRC_Purchase_Order_Register';
            case 'GRN':
                return 'PRC_GRN_Register';
            default:
                return '';
        }
    };

    const getTitle = (): string => {
        switch (recordType) {
            case 'MR':
                return 'Material Requests';
            case 'PR':
                return 'Purchase Requisitions';
            case 'PO':
                return 'Purchase Orders';
            case 'GRN':
                return 'Goods Received Notes';
            default:
                return '';
        }
    };

    const getNewButtonText = (): string => {
        switch (recordType) {
            case 'MR':
                return '+ New Request';
            case 'PR':
                return '+ New Requisition';
            case 'PO':
                return '+ New Order';
            case 'GRN':
                return '+ New GRN';
            default:
                return '+ New Record';
        }
    };

    const getColumns = (): IDataGridColumn[] => {
        switch (recordType) {
            case 'MR':
                return [
                    { key: 'Title', name: 'Request #', fieldName: 'Title', minWidth: 120, isResizable: true },
                    { key: 'Project_Code', name: 'Project Code', fieldName: 'Project_Code', minWidth: 120, isResizable: true },
                    { key: 'Material', name: 'Material', fieldName: 'Material', minWidth: 150, isResizable: true },
                    { key: 'Quantity', name: 'Quantity', fieldName: 'Quantity', minWidth: 100, isResizable: true },
                    { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, isResizable: true },
                ];
            case 'PR':
                return [
                    { key: 'Title', name: 'Requisition #', fieldName: 'Title', minWidth: 120, isResizable: true },
                    { key: 'Project_Code', name: 'Project Code', fieldName: 'Project_Code', minWidth: 120, isResizable: true },
                    { key: 'Description', name: 'Description', fieldName: 'Description', minWidth: 150, isResizable: true },
                    { key: 'Quantity', name: 'Quantity', fieldName: 'Quantity', minWidth: 100, isResizable: true },
                    { key: 'Estimated_Cost', name: 'Est. Cost', fieldName: 'Estimated_Cost', minWidth: 120, isResizable: true },
                    { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, isResizable: true },
                ];
            case 'PO':
                return [
                    { key: 'Title', name: 'PO #', fieldName: 'Title', minWidth: 120, isResizable: true },
                    { key: 'Vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150, isResizable: true },
                    { key: 'TotalAmount', name: 'Total Amount', fieldName: 'TotalAmount', minWidth: 120, isResizable: true },
                    { key: 'Delivery_Date', name: 'Delivery Date', fieldName: 'Delivery_Date', minWidth: 120, isResizable: true },
                    { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, isResizable: true },
                ];
            case 'GRN':
                return [
                    { key: 'Title', name: 'GRN #', fieldName: 'Title', minWidth: 120, isResizable: true },
                    { key: 'PO_Number', name: 'PO Number', fieldName: 'PO_Number', minWidth: 120, isResizable: true },
                    { key: 'Vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150, isResizable: true },
                    { key: 'Quantity_Received', name: 'Qty Received', fieldName: 'Quantity_Received', minWidth: 120, isResizable: true },
                    { key: 'Received_Date', name: 'Received Date', fieldName: 'Received_Date', minWidth: 120, isResizable: true },
                ];
            default:
                return [];
        }
    };



    const handleRefresh = (): void => {
        setRefreshKey((prev) => prev + 1);
    };

    const handleNewRecord = (): void => {
        setIsFormPanelOpen(true);
    };

    const handleFormSubmit = (): void => {
        setIsFormPanelOpen(false);
        handleRefresh();
    };

    const handleRowClick = (record: IProcurementRecord): void => {
        setSelectedRecord(record);
        setIsDetailsPanelOpen(true);
    };

    const closeDetailsPanel = (): void => {
        setIsDetailsPanelOpen(false);
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

    const renderForm = (): React.ReactNode => {
        switch (recordType) {
            case 'MR':
                return (
                    <MaterialRequestForm
                        isOpen={isFormPanelOpen}
                        onDismiss={() => setIsFormPanelOpen(false)}
                        onSubmitSuccess={handleFormSubmit}
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        webPartContext={webPartContext}
                    />
                );
            case 'PR':
                return (
                    <PurchaseRequisitionForm
                        isOpen={isFormPanelOpen}
                        onDismiss={() => setIsFormPanelOpen(false)}
                        onSubmitSuccess={handleFormSubmit}
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        webPartContext={webPartContext}
                    />
                );
            case 'PO':
                return (
                     <PurchaseOrderForm
                         isOpen={isFormPanelOpen}
                         onDismiss={() => setIsFormPanelOpen(false)}
                         onSubmitSuccess={handleFormSubmit}
                         spHttpClient={spHttpClient}
                         pageContext={pageContext}
                         webPartContext={webPartContext}
                     />
                );
            case 'GRN':
                return (
                    <GoodsReceivedNoteForm
                        isOpen={isFormPanelOpen}
                        onDismiss={() => setIsFormPanelOpen(false)}
                        onSubmitSuccess={handleFormSubmit}
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        webPartContext={webPartContext}
                    />
                );
            default:
                return null;
        }
    };

    return (
        <div className={classNames.root}>
            {/* Header Section */}
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        {getTitle()}
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Manage and track {getTitle().toLowerCase()}
                    </Text>
                </div>
                <div className={classNames.headerActions}>
                    <PrimaryButton
                        text={getNewButtonText()}
                        onClick={handleNewRecord}
                        iconProps={{ iconName: 'Add' }}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={handleRefresh}
                        title="Refresh"
                        ariaLabel="Refresh"
                    />
                </div>
            </div>

            {/* Grid Container */}
            <div className={classNames.gridContainer}>
                <DataGrid
                    key={`procurement-${recordType}-${refreshKey}`}
                    listName={getListName()}
                    columns={getColumns()}
                      expandQuery={
                        recordType === 'MR' ? 'Project_Code,Material' :
                        recordType === 'PR' ? 'Project_Code' :
                        recordType === 'PO' ? 'Vendor,Project_Code' :
                        recordType === 'GRN' ? 'Vendor,PO_Number' :
                        undefined
                      }
                    pageSize={20}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onRowSelected={handleRowClick}
                />
            </div>

            {/* Form Panel */}
            {renderForm()}

            {/* Details Panel - Professional view like ProjectDetailsPanel */}
            {selectedRecord && isDetailsPanelOpen && (
                <ProcurementDetailsPanel
                    isOpen={isDetailsPanelOpen}
                    record={selectedRecord}
                    recordType={recordType}
                    onDismiss={closeDetailsPanel}
                    onRefresh={handleRefresh}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    showApprovalButtons={true}
                    currentUserDisplayName={userDisplayName}
                    webPartContext={webPartContext}
                    onApprove={() => {
                        closeDetailsPanel();
                    }}
                    onReject={() => {
                        closeDetailsPanel();
                    }}
                     onTrackApproval={openApprovalTracker}
                 />
            )}

            {/* Approval Tracker Panel */}
            {selectedRecord && isApprovalTrackerOpen && (
                <ApprovalTrackerPanel
                    isOpen={isApprovalTrackerOpen}
                    recordId={selectedRecord.ID}
                    listName={getListName()}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onDismiss={closeApprovalTracker}
                    onRefresh={handleApprovalRefresh}
                />
            )}

            {/* Approval Tracker Panel */}
            {selectedRecord && isApprovalTrackerOpen && (
                <ApprovalTrackerPanel
                    isOpen={isApprovalTrackerOpen}
                    recordId={selectedRecord.ID}
                    listName={getListName()}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onDismiss={closeApprovalTracker}
                    onRefresh={handleApprovalRefresh}
                />
            )}
        </div>
    );
};

export default ProcurementSubModule;
