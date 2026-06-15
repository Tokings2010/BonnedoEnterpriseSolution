import * as React from 'react';
import { getTheme, mergeStyleSets, Pivot, PivotItem } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

import ReportsProcurement from './ReportsProcurement';
import ReportsFinance from './ReportsFinance';
import ReportsProjects from './ReportsProjects';

export interface IReportsModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  onBack?: () => void;
}

const ReportsModule: React.FC<IReportsModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  onBack,
}) => {
  const theme = getTheme();
  const [selectedTab, setSelectedTab] = React.useState<string>('procurement');
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);

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

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item?: PivotItem) => {
          if (item) {
            setSelectedTab(item.props.itemKey || 'procurement');
          }
        }}
        className={classNames.pivotContainer}
      >
        <PivotItem
          itemKey="procurement"
          headerText="Procurement Reports"
          itemIcon="ShoppingCart"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <ReportsProcurement
              key={`procurement-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="finance"
          headerText="Finance Reports"
          itemIcon="Money"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <ReportsFinance
              key={`finance-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="projects"
          headerText="Project Reports"
          itemIcon="ProjectCollection"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <ReportsProjects
              key={`projects-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default ReportsModule;
