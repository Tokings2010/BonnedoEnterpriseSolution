import * as React from 'react';
import { getTheme, mergeStyleSets, Pivot, PivotItem, Text } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IUserPermissions, MaterialSubModuleKey } from '../../models/PermissionModels';

import MaterialsModule from '../MaterialsModule';
import InventoryTab from './InventoryTab';
import MovementsTab from './MovementsTab';
import QRScanMovePanel, { IQRScanRequestContext } from './QRScanMovePanel';
import WarehousesModule from '../WarehousesModule';

export interface IMaterialsManagementModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  userPermissions?: IUserPermissions;
}

interface IMaterialTabConfig {
  key: MaterialSubModuleKey;
  headerText: string;
  itemIcon: string;
}

const MATERIAL_TABS: IMaterialTabConfig[] = [
  { key: 'material-master', headerText: 'Material Master', itemIcon: 'Package' },
  { key: 'inventory', headerText: 'Inventory', itemIcon: 'BoxMultipleSolid' },
  { key: 'qr-scan', headerText: 'QR Scan & Move', itemIcon: 'QRCode' },
  { key: 'movements', headerText: 'Movements', itemIcon: 'Switch' },
  { key: 'warehouse', headerText: 'Warehouses', itemIcon: 'CityNext' },
];

const MaterialsManagementModule: React.FC<IMaterialsManagementModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  userPermissions,
}) => {
  const theme = getTheme();
  const [selectedTab, setSelectedTab] = React.useState<MaterialSubModuleKey>('material-master');
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);
  const [qrScanContext, setQrScanContext] = React.useState<IQRScanRequestContext | undefined>(undefined);

  const visibleTabs = React.useMemo(() => {
    if (!userPermissions) {
      return MATERIAL_TABS;
    }
    return MATERIAL_TABS.filter((tab) => userPermissions.hasSubModulePermission(tab.key));
  }, [userPermissions]);

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
    emptyTabs: {
      padding: '40px',
      textAlign: 'center',
      color: theme.palette.neutralSecondary,
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

  React.useEffect(() => {
    if (visibleTabs.length > 0 && !visibleTabs.some((tab) => tab.key === selectedTab)) {
      setSelectedTab(visibleTabs[0].key);
    }
  }, [visibleTabs, selectedTab]);

  const handleRefresh = (): void => {
    setRefreshKey((prev) => prev + 1);
  };

  const handleQrScanRequested = (context?: IQRScanRequestContext): void => {
    setQrScanContext(context);
    setSelectedTab('qr-scan');
  };

  const handleScanComplete = (): void => {
    window.dispatchEvent(new CustomEvent('refreshData'));
  };

  const renderTabContent = (tabKey: MaterialSubModuleKey): React.ReactNode => {
    switch (tabKey) {
      case 'material-master':
        return (
          <MaterialsModule
            key={`material-master-${refreshKey}`}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
          />
        );
      case 'inventory':
        return (
          <InventoryTab
            key={`inventory-${refreshKey}`}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            isMobileView={isMobileView}
            onRefresh={handleRefresh}
            onQrScanRequested={handleQrScanRequested}
          />
        );
      case 'qr-scan':
        return (
          <QRScanMovePanel
            key={`qr-scan-${refreshKey}`}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            context={qrScanContext}
            onScanComplete={handleScanComplete}
          />
        );
      case 'movements':
        return (
          <MovementsTab
            key={`movements-${refreshKey}`}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            isMobileView={isMobileView}
            onRefresh={handleRefresh}
            onQrScanRequested={handleQrScanRequested}
          />
        );
      case 'warehouse':
        return (
          <WarehousesModule
            key={`warehouses-${refreshKey}`}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
          />
        );
      default:
        return null;
    }
  };

  if (visibleTabs.length === 0) {
    return (
      <div className={classNames.emptyTabs}>
        <Text variant="large">No Material sub-modules available for your role.</Text>
      </div>
    );
  }

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item?: PivotItem) => {
          if (item?.props.itemKey) {
            const nextTab = item.props.itemKey as MaterialSubModuleKey;
            setQrScanContext(nextTab === 'qr-scan' ? undefined : qrScanContext);
            setSelectedTab(nextTab);
          }
        }}
        className={classNames.pivotContainer}
      >
        {visibleTabs.map((tab) => (
          <PivotItem
            key={tab.key}
            itemKey={tab.key}
            headerText={tab.headerText}
            itemIcon={tab.itemIcon}
          >
            <div className={classNames.tabContent}>{renderTabContent(tab.key)}</div>
          </PivotItem>
        ))}
      </Pivot>
    </div>
  );
};

export default MaterialsManagementModule;
