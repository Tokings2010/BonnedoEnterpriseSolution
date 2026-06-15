import * as React from 'react';
import { getTheme, mergeStyleSets, Pivot, PivotItem } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import MaterialsModule from './MaterialsModule';
import VendorsModule from './VendorsModule';

export interface IMasterDataModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
}

const MasterDataModule: React.FC<IMasterDataModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
}) => {
  const theme = getTheme();
  const [selectedTab, setSelectedTab] = React.useState<'materials' | 'vendors'>('materials');

  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      backgroundColor: theme.palette.white,
      padding: '16px',
    },
    pivot: {
      flex: 1,
      display: 'flex',
      flexDirection: 'column',
      overflow: 'hidden',
    },
  });

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item) => {
          if (item?.props.itemKey) {
            setSelectedTab(item.props.itemKey as 'materials' | 'vendors');
          }
        }}
        className={classNames.pivot}
      >
        <PivotItem headerText="Materials" itemKey="materials" style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ height: '100%', overflow: 'hidden' }}>
            <MaterialsModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
            />
          </div>
        </PivotItem>
        <PivotItem headerText="Vendors" itemKey="vendors" style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ height: '100%', overflow: 'hidden' }}>
            <VendorsModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
            />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default MasterDataModule;
