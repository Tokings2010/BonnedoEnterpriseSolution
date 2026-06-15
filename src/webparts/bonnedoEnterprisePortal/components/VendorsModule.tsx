import * as React from 'react';
import {
    getTheme,
    mergeStyleSets,
    PrimaryButton,
    IconButton,
    Text,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import DataGrid from './DataGrid';
import { IDataGridColumn } from './DataGrid';
import VendorForm from './VendorForm';
import VendorDetailsPanel from './VendorDetailsPanel';

export interface IVendorsModuleProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    userDisplayName: string;
}

export interface IVendor {
    ID: number;
    Vendor_Code: string;
    Vendor_Name: string;
    Vendor_Category: string;
    Vendor_Status: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    [key: string]: any;
}

const VendorsModule: React.FC<IVendorsModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
}) => {
  const theme = getTheme();
  const [selectedVendor, setSelectedVendor] = React.useState<IVendor | undefined>(undefined);
  const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
  const [isFormPanelOpen, setIsFormPanelOpen] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);

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
    });

    const vendorColumns: IDataGridColumn[] = [
        {
            key: 'Vendor_Code',
            name: 'Vendor Code',
            fieldName: 'Vendor_Code',
            minWidth: 120,
            isResizable: true,
        },
        {
            key: 'Vendor_Name',
            name: 'Vendor Name',
            fieldName: 'Vendor_Name',
            minWidth: 180,
            isResizable: true,
        },
        {
            key: 'Vendor_Category',
            name: 'Category',
            fieldName: 'Vendor_Category',
            minWidth: 150,
            isResizable: true,
        },
        {
            key: 'Vendor_Status',
            name: 'Status',
            fieldName: 'Vendor_Status',
            minWidth: 100,
            isResizable: true,
            onRender: (item: IVendor) => {
                let color = '';
                switch (item.Vendor_Status) {
                    case 'Active':
                        color = theme.palette.green;
                        break;
                    case 'Inactive':
                        color = theme.palette.red;
                        break;
                    case 'Pending':
                        color = theme.palette.orange;
                        break;
                    case 'Blocked':
                        color = theme.palette.red;
                        break;
                    default:
                        color = theme.palette.neutralSecondary;
                }
                return <span style={{ color }}>{item.Vendor_Status}</span>;
            },
        },
    ];

    const handleRefresh = (): void => {
        setRefreshKey((prev) => prev + 1);
    };

    const handleNewVendor = (): void => {
        setIsFormPanelOpen(true);
    };

    const handleFormSubmit = (): void => {
        setIsFormPanelOpen(false);
        handleRefresh();
    };

  const handleRowSelected = (vendor: IVendor): void => {
    setSelectedVendor(vendor);
    setIsDetailsPanelOpen(true);
  };

  const closeDetailsPanel = (): void => {
    setIsDetailsPanelOpen(false);
    setSelectedVendor(undefined);
  };



    return (
        <div className={classNames.root}>
            {/* Header Section */}
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Vendors Master
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Manage and track all vendors
                    </Text>
                </div>
                <div className={classNames.headerActions}>
                    <PrimaryButton
                        text="+ Add Vendor"
                        onClick={handleNewVendor}
                        iconProps={{ iconName: 'Add' }}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={handleRefresh}
                        title="Refresh vendors"
                        ariaLabel="Refresh vendors"
                    />
                </div>
            </div>

            {/* Grid Container */}
            <div className={classNames.gridContainer}>
                <DataGrid
                    key={`vendors-${refreshKey}`}
                    listName="ENT_Vendors_Master"
                    columns={vendorColumns}
                    pageSize={20}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    onRowSelected={handleRowSelected}
                />
            </div>

            {/* Vendor Form Panel */}
            <VendorForm
                isOpen={isFormPanelOpen}
                onDismiss={() => setIsFormPanelOpen(false)}
                onSubmitSuccess={handleFormSubmit}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
            />

            {/* Vendor Details Panel */}
            <VendorDetailsPanel
                isOpen={isDetailsPanelOpen}
                vendor={selectedVendor}
                onDismiss={closeDetailsPanel}
                onRefresh={handleRefresh}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
            />
        </div>
    );
};

export default VendorsModule;
