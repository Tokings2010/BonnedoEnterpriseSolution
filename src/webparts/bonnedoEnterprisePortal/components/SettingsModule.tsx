import * as React from 'react';
import { getTheme, mergeStyleSets, Pivot, PivotItem } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import SettingsApprovalMatrix from './SettingsApprovalMatrix';
import SettingsSystemConfig from './SettingsSystemConfig';
import SettingsUserRoles from './SettingsUserRoles';
import SettingsNotifications from './SettingsNotifications';
import ApprovalMatrixForm from './ApprovalMatrixForm';
import UserRoleForm from './UserRoleForm';

export interface ISettingsModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  webPartContext?: WebPartContext;
  onBack?: () => void;
}

const SettingsModule: React.FC<ISettingsModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  webPartContext,
  onBack,
}) => {
  const theme = getTheme();
  const [selectedTab, setSelectedTab] = React.useState<string>('approval-matrix');
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);

  // Form panel states
  const [isApprovalMatrixFormOpen, setIsApprovalMatrixFormOpen] = React.useState(false);
  const [isUserRoleFormOpen, setIsUserRoleFormOpen] = React.useState(false);

  // Handle window resize for responsive behavior
  React.useEffect(() => {
    const handleResize = (): void => {
      setIsMobileView(window.innerWidth < 768);
    };

    window.addEventListener('resize', handleResize);
    return (): void => {
      window.removeEventListener('resize', handleResize);
    };
  }, []);

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
  });

  const handleRefresh = (): void => {
    setRefreshKey((prev) => prev + 1);
  };

  const handleApprovalMatrixFormSubmit = (): void => {
    setIsApprovalMatrixFormOpen(false);
    handleRefresh();
  };

  const handleUserRoleFormSubmit = (): void => {
    setIsUserRoleFormOpen(false);
    handleRefresh();
  };

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item?: PivotItem) => {
          if (item) {
            setSelectedTab(item.props.itemKey || 'approval-matrix');
          }
        }}
        className={classNames.pivotContainer}
      >
        <PivotItem
          itemKey="approval-matrix"
          headerText="Approval Matrix"
          itemIcon="CheckMark"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <SettingsApprovalMatrix
              key={`approval-matrix-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              webPartContext={webPartContext}
              isMobileView={isMobileView}
              onNewApproval={() => setIsApprovalMatrixFormOpen(true)}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="system-config"
          headerText="System Configuration"
          itemIcon="Settings"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <SettingsSystemConfig
              key={`system-config-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="user-roles"
          headerText="User Roles"
          itemIcon="People"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <SettingsUserRoles
              key={`user-roles-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onNewUserRole={() => setIsUserRoleFormOpen(true)}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="notifications"
          headerText="Notification Settings"
          itemIcon="Mail"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <SettingsNotifications
              key={`notifications-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>
      </Pivot>

      {/* Approval Matrix Form Panel */}
      <ApprovalMatrixForm
        isOpen={isApprovalMatrixFormOpen}
        onDismiss={() => setIsApprovalMatrixFormOpen(false)}
        onSubmitSuccess={handleApprovalMatrixFormSubmit}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
      />

      {/* User Role Form Panel */}
      <UserRoleForm
        isOpen={isUserRoleFormOpen}
        onDismiss={() => setIsUserRoleFormOpen(false)}
        onSubmitSuccess={handleUserRoleFormSubmit}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
        editMode={false}
      />
    </div>
  );
};

export default SettingsModule;
