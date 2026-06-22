import * as React from 'react';
import {
    getTheme,
    mergeStyleSets,
    Pivot,
    PivotItem,
    PrimaryButton,
    Stack,
    Text,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ProcurementSubModule from './ProcurementSubModule';
import BidManagementTab from './ProcurementModule/BidManagementTab';
import NewProcurementRequestPanel from './ProcurementModule/NewProcurementRequestPanel';
import VendorRegistrationTab from './ProcurementModule/VendorRegistrationTab';

export interface IProcurementModuleProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    userDisplayName: string;
    onBack?: () => void;
    webPartContext?: WebPartContext;
}

type ProcurementTabKey = 'REQUEST' | 'BIDS' | 'VENDORS' | 'MR' | 'PR' | 'PO' | 'GRN';

interface IProcurementTabConfig {
    key: ProcurementTabKey;
    headerText: string;
    itemIcon: string;
}

const PROCUREMENT_TABS: IProcurementTabConfig[] = [
    { key: 'REQUEST', headerText: 'New Procurement Request', itemIcon: 'Add' },
    { key: 'BIDS', headerText: 'Bid Management', itemIcon: 'CompareUneven' },
    { key: 'VENDORS', headerText: 'Vendor Registration', itemIcon: 'AddFriend' },
    { key: 'MR', headerText: 'Material Request', itemIcon: 'ReceiptProcessing' },
    { key: 'PR', headerText: 'Purchase Requisition', itemIcon: 'ClipboardList' },
    { key: 'PO', headerText: 'Purchase Order', itemIcon: 'OrderStatus' },
    { key: 'GRN', headerText: 'Goods Received Note', itemIcon: 'Package' },
];

const ProcurementModule: React.FC<IProcurementModuleProps> = ({
    spHttpClient,
    pageContext,
    userDisplayName,
    webPartContext,
}) => {
    const theme = getTheme();
    const [selectedTab, setSelectedTab] = React.useState<ProcurementTabKey>('REQUEST');
    const [refreshKey, setRefreshKey] = React.useState(0);
    const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);
    const [isRequestPanelOpen, setIsRequestPanelOpen] = React.useState(false);

    const classNames = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            backgroundColor: theme.palette.white,
        },
        pivotContainer: {
            flex: 1,
            overflow: 'auto',
        },
        tabContent: {
            padding: '16px',
            height: '100%',
        },
        requestLauncher: {
            minHeight: 'calc(100vh - 220px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
        },
        requestCard: {
            width: '100%',
            maxWidth: '640px',
            padding: '28px',
            borderRadius: '12px',
            border: `1px solid ${theme.palette.neutralLight}`,
            backgroundColor: theme.palette.neutralLighterAlt,
            textAlign: 'center',
        },
    });

    React.useEffect(() => {
        const handleResize = (): void => {
            setIsMobileView(window.innerWidth < 768);
        };

        window.addEventListener('resize', handleResize);
        return (): void => {
            window.removeEventListener('resize', handleResize);
        };
    }, []);

    React.useEffect(() => {
        const handleRefresh = (): void => {
            setRefreshKey((prev) => prev + 1);
        };

        window.addEventListener('refreshData', handleRefresh);
        return (): void => {
            window.removeEventListener('refreshData', handleRefresh);
        };
    }, []);

    const handleRefresh = (): void => {
        setRefreshKey((prev) => prev + 1);
    };

    const renderProcurementSubModule = (recordType: 'MR' | 'PR' | 'PO' | 'GRN'): React.ReactNode => (
        <ProcurementSubModule
            key={`procurement-${recordType}-${refreshKey}`}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            recordType={recordType}
            webPartContext={webPartContext}
        />
    );

    const renderTabContent = (tabKey: ProcurementTabKey): React.ReactNode => {
        switch (tabKey) {
            case 'REQUEST':
                return (
                    <div className={classNames.requestLauncher}>
                        <div className={classNames.requestCard}>
                            <Text variant="xLarge" block style={{ fontWeight: 600, marginBottom: '8px' }}>
                                New Procurement Request
                            </Text>
                            <Text variant="medium" block style={{ color: theme.palette.neutralSecondary, marginBottom: '20px' }}>
                                Create a procurement request, add line items, and choose the sourcing method.
                            </Text>
                            <Stack horizontal horizontalAlign="center">
                                <PrimaryButton
                                    text="Create Request"
                                    onClick={() => setIsRequestPanelOpen(true)}
                                    iconProps={{ iconName: 'Add' }}
                                />
                            </Stack>
                        </div>
                    </div>
                );
            case 'BIDS':
                return (
                    <BidManagementTab
                        key={`bid-management-${refreshKey}`}
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        isMobileView={isMobileView}
                        onRefresh={handleRefresh}
                    />
                );
            case 'VENDORS':
                return (
                    <VendorRegistrationTab
                        key={`vendor-registration-${refreshKey}`}
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        isMobileView={isMobileView}
                        onRefresh={handleRefresh}
                    />
                );
            case 'MR':
                return renderProcurementSubModule('MR');
            case 'PR':
                return renderProcurementSubModule('PR');
            case 'PO':
                return renderProcurementSubModule('PO');
            case 'GRN':
                return renderProcurementSubModule('GRN');
            default:
                return null;
        }
    };

    return (
        <div className={classNames.root}>
            <Pivot
                selectedKey={selectedTab}
                onLinkClick={(item?: PivotItem) => {
                    if (item?.props.itemKey) {
                        setSelectedTab(item.props.itemKey as ProcurementTabKey);
                    }
                }}
                className={classNames.pivotContainer}
            >
                {PROCUREMENT_TABS.map((tab) => (
                    <PivotItem
                        key={tab.key}
                        itemKey={tab.key}
                        headerText={tab.headerText}
                        itemIcon={tab.itemIcon}
                    >
                        <div className={classNames.tabContent}>
                            {renderTabContent(tab.key)}
                        </div>
                    </PivotItem>
                ))}
            </Pivot>

            <NewProcurementRequestPanel
                isOpen={isRequestPanelOpen}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
                onDismiss={() => setIsRequestPanelOpen(false)}
                onRequestCreated={() => {
                    setIsRequestPanelOpen(false);
                    handleRefresh();
                }}
            />
        </div>
    );
};

export default ProcurementModule;
