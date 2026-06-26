import * as React from 'react';
import {
  Text, getTheme, mergeStyleSets, IconButton, Spinner, SpinnerSize,
  Dropdown, IDropdownOption,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import EnhancedDataGrid from './EnhancedDataGrid';
import { IDataGridColumn } from './EnhancedDataGrid';

export interface IPurchaseOrderTrackerTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
}

const PurchaseOrderTrackerTab: React.FC<IPurchaseOrderTrackerTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
}) => {
  const theme = getTheme();
  const [loading, setLoading] = React.useState(true);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [projectFilter, setProjectFilter] = React.useState<string>('All');
  const [projectOptions, setProjectOptions] = React.useState<IDropdownOption[]>([]);

  const classNames = mergeStyleSets({
    container: { padding: '20px', height: '100%', display: 'flex', flexDirection: 'column' },
    header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' as const },
    headerTitle: { flex: 1, minWidth: '200px' },
    headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' },
    gridContainer: { flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' },
    loadingContainer: { display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' },
  });

  const loadProjects = React.useCallback(async (): Promise<void> => {
    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('ENT_Project_Master')/items?$select=Project_Code&$top=200`;
      const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json();
        const opts: IDropdownOption[] = (data.value || [])
          .map((p: any) => p.Project_Code).filter(Boolean)
          .map((code: string) => ({ key: code, text: code }));
        setProjectOptions([{ key: 'All', text: 'All Projects' }, ...opts]);
      }
    } catch { /* silent */ }
    setLoading(false);
  }, [spHttpClient, pageContext]);

  React.useEffect(() => { loadProjects().catch(() => undefined); }, [loadProjects]);

  const handleRefresh = (): void => { setRefreshKey(prev => prev + 1); onRefresh(); };

  const poColumns: IDataGridColumn[] = [
    { key: 'PO_Number', name: 'PO Number', fieldName: 'Title', minWidth: 120, isResizable: true,
      onRender: (item: any) => <span style={{ fontWeight: 600, fontFamily: "'Cascadia Code', Consolas, monospace", fontSize: 12 }}>{item.PO_Number || item.Title}</span> },
    { key: 'Project_Code', name: 'Project', fieldName: 'Project_Code', minWidth: 100, isResizable: true,
      onRender: (item: any) => {
        const v = item.Project_Code || item.Project_x0020_Code || '';
        return <span>{typeof v === 'object' ? (v.Title || '') : v}</span>;
      }},
    { key: 'Vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150, isResizable: true,
      onRender: (item: any) => <span>{typeof item.Vendor === 'object' ? (item.Vendor.Title || '') : (item.Vendor || '')}</span> },
    { key: 'TotalAmount', name: 'Amount', fieldName: 'TotalAmount', minWidth: 120, isResizable: true,
      onRender: (item: any) => <span style={{ fontWeight: 600 }}>{`\u20A6${(item.TotalAmount || item.Amount || 0).toLocaleString()}`}</span> },
    { key: 'Amount_Paid', name: 'Paid', fieldName: 'Amount_Paid', minWidth: 110, isResizable: true,
      onRender: (item: any) => <span>{`\u20A6${(item.Amount_Paid || 0).toLocaleString()}`}</span> },
    { key: 'Balance', name: 'Balance', fieldName: '', minWidth: 110, isResizable: true,
      onRender: (item: any) => {
        const balance = (item.TotalAmount || item.Amount || 0) - (item.Amount_Paid || 0);
        return <span style={{ fontWeight: 600, color: balance > 0 ? '#D13438' : '#107C10' }}>{`\u20A6${balance.toLocaleString()}`}</span>;
      }},
    { key: 'Approval_Status', name: 'Status', fieldName: 'Approval_Status', minWidth: 100, isResizable: true,
      onRender: (item: any) => <span>{item.Approval_Status || item.Status || ''}</span> },
    { key: 'DeliveryDate', name: 'Delivery', fieldName: 'DeliveryDate', minWidth: 100, isResizable: true,
      onRender: (item: any) => { const d = item.DeliveryDate || item.Delivery_x0020_Date || ''; return <span>{d ? new Date(d).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : ''}</span>; }},
  ];

  return (
    <div className={classNames.container}>
      <div className={classNames.header}>
        <div className={classNames.headerTitle}>
          <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
            PO Tracker
          </Text>
          <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
            Purchase order payment tracking and balances
          </Text>
        </div>
        <div className={classNames.headerActions}>
          <Dropdown
            placeholder="Filter by project"
            options={projectOptions}
            selectedKey={projectFilter}
            onChange={(_, opt) => setProjectFilter(opt?.key as string || 'All')}
            styles={{ root: { minWidth: '180px' } }}
          />
          <IconButton iconProps={{ iconName: 'Refresh' }} onClick={handleRefresh} title="Refresh" />
        </div>
      </div>

      <div className={classNames.gridContainer}>
        <EnhancedDataGrid
          key={`po-tracker-${refreshKey}`}
          listName="PRC_Purchase_Order_Register"
          columns={poColumns}
          pageSize={20}
          spHttpClient={spHttpClient}
          pageContext={pageContext}
          expandQuery="Vendor"
          showExport
        />
      </div>
    </div>
  );
};

export default PurchaseOrderTrackerTab;
