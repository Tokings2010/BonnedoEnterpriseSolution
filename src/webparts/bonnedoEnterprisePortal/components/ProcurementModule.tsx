import * as React from 'react';
import { getTheme, mergeStyleSets, Pivot, PivotItem } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ProcurementSubModule from './ProcurementSubModule';

export interface IProcurementModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  onBack?: () => void;
  webPartContext?: WebPartContext;
}

const ProcurementModule: React.FC<IProcurementModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  onBack,
  webPartContext,
}) => {
  const theme = getTheme();
  const [selectedTab, setSelectedTab] = React.useState<'MR' | 'PR' | 'PO' | 'GRN'>('MR');

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

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item) => {
          if (item?.props.itemKey) {
            setSelectedTab(item.props.itemKey as 'MR' | 'PR' | 'PO' | 'GRN');
          }
        }}
        className={classNames.pivotContainer}
      >
        <PivotItem headerText="Material Request" itemKey="MR">
          <div style={{ padding: '16px', height: '100%' }}>
            <ProcurementSubModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
              recordType="MR"
              webPartContext={webPartContext}
            />
          </div>
        </PivotItem>

        <PivotItem headerText="Purchase Requisition" itemKey="PR">
          <div style={{ padding: '16px', height: '100%' }}>
            <ProcurementSubModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
              recordType="PR"
              webPartContext={webPartContext}
            />
          </div>
        </PivotItem>

        <PivotItem headerText="Purchase Order" itemKey="PO">
          <div style={{ padding: '16px', height: '100%' }}>
            <ProcurementSubModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
              recordType="PO"
              webPartContext={webPartContext}
            />
          </div>
        </PivotItem>

        <PivotItem headerText="Goods Received Note" itemKey="GRN">
          <div style={{ padding: '16px', height: '100%' }}>
            <ProcurementSubModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
              recordType="GRN"
              webPartContext={webPartContext}
            />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default ProcurementModule;
