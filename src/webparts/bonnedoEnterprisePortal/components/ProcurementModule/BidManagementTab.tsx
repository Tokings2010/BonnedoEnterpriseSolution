import * as React from 'react';
import {
  mergeStyleSets,
  Text, Icon,
  SearchBox, Dropdown, IDropdownOption,
  CommandBar, ICommandBarItemProps,
  Spinner, SpinnerSize,
  MessageBar, MessageBarType,
  Stack,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import BidDetailPanel from './BidDetailPanel';
import NewProcurementRequestPanel from './NewProcurementRequestPanel';
import EnhancedDataGrid from '../EnhancedDataGrid';
import { IDataGridColumn } from '../EnhancedDataGrid';
import { Tag, formatCurrency, formatDate } from '../TagRenderer';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IBidManagementTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
}

export interface IProcurementRequest {
  ID: number;
  RequestID: string;
  Title: string;
  Project_Code: string;
  Requested_By: string;
  Request_Date: string;
  Required_Date: string;
  Sourcing_Method: string;
  Assigned_Vendor: string;
  Status: string;
  Budget_Estimate: number;
  Category: string;
  Justification: string;
  Approved_By: string | undefined;
  Approval_Date: string | undefined;
}

interface ISharePointListItem {
  ID: number;
  [key: string]: unknown;
}

const getString = (value: unknown): string => typeof value === 'string' ? value : '';
const getOptionalString = (value: unknown): string | undefined => typeof value === 'string' && value ? value : undefined;
const getNumber = (value: unknown): number => typeof value === 'number' ? value : 0;

const getPersonTitle = (value: unknown): string | undefined => {
  if (value && typeof value === 'object' && 'Title' in value) {
    const title = (value as { Title?: unknown }).Title;
    return typeof title === 'string' ? title : undefined;
  }

  return undefined;
};

// ─── Status Color Map ────────────────────────────────────────────────────────

const REQUEST_STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  Draft: { bg: '#F3F4F6', text: '#374151' },
  'Pending Approval': { bg: '#FEF3C7', text: '#92400E' },
  Approved: { bg: '#D1FAE5', text: '#065F46' },
  'Bid In Progress': { bg: '#DBEAFE', text: '#1E40AF' },
  'Quotes Received': { bg: '#EDE9FE', text: '#5B21B6' },
  Awarded: { bg: '#D1FAE5', text: '#065F46' },
  Cancelled: { bg: '#FEE2E2', text: '#991B1B' },
};

const SOURCING_COLORS: Record<string, { bg: string; text: string }> = {
  'Single Source': { bg: '#EDE9FE', text: '#7C3AED' },
  'Competitive Bid': { bg: '#DBEAFE', text: '#2563EB' },
};

const STATUS_FILTER_OPTIONS: IDropdownOption[] = [
  { key: 'All', text: 'All Statuses' },
  { key: 'Draft', text: 'Draft' },
  { key: 'Pending Approval', text: 'Pending Approval' },
  { key: 'Approved', text: 'Approved' },
  { key: 'Bid In Progress', text: 'Bid In Progress' },
  { key: 'Quotes Received', text: 'Quotes Received' },
  { key: 'Awarded', text: 'Awarded' },
  { key: 'Cancelled', text: 'Cancelled' },
];

// ─── Styles ──────────────────────────────────────────────────────────────────

const getStyles = (): ReturnType<typeof mergeStyleSets> =>
  mergeStyleSets({
    container: {
      display: 'flex',
      flexDirection: 'column',
      gap: '16px',
    },
    headerRow: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      flexWrap: 'wrap',
      gap: '12px',
    },
    filterRow: {
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
      flexWrap: 'wrap',
    },
    statsRow: {
      display: 'flex',
      gap: '16px',
      flexWrap: 'wrap',
    },
    statCard: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '12px 20px',
      display: 'flex',
      alignItems: 'center',
      gap: '10px',
    },
    statValue: {
      fontSize: '20px',
      fontWeight: '700',
      color: '#1E2532',
    },
    statLabel: {
      fontSize: '12px',
      color: '#5A6A85',
    },
    statusPill: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
    },
    gridContainer: { flex: 1, display: 'flex', flexDirection: 'column' },
    emptyState: {
      textAlign: 'center',
      padding: '60px 20px',
      color: '#5A6A85',
    },
  });

// ─── Component ───────────────────────────────────────────────────────────────

const BidManagementTab: React.FC<IBidManagementTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
}) => {
  const [items, setItems] = React.useState<IProcurementRequest[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IProcurementRequest[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [searchText, setSearchText] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<string>('All');
  const [selectedRequest, setSelectedRequest] = React.useState<IProcurementRequest | undefined>(undefined);
  const [isDetailPanelOpen, setIsDetailPanelOpen] = React.useState(false);
  const [isNewRequestPanelOpen, setIsNewRequestPanelOpen] = React.useState(false);

  const classNames = getStyles() as unknown as { [key: string]: string };

  // ─── Data Fetching ─────────────────────────────────────────────────────────

  const fetchRequests = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(undefined);
    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Procurement_Requests')/items?$top=500&$orderby=Created desc`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch procurement requests: ${response.status}`);
      }

      const data = await response.json();
      const requests: IProcurementRequest[] = (data.value || []).map((item: ISharePointListItem) => ({
        ID: item.ID,
        RequestID: getString(item.RequestID),
        Title: getString(item.Title),
        Project_Code: getString(item.Project_Code),
        Requested_By: getPersonTitle(item.Requested_By) || '',
        Request_Date: getString(item.Request_Date),
        Required_Date: getString(item.Required_Date),
        Sourcing_Method: getString(item.Sourcing_Method),
        Assigned_Vendor: getString(item.Assigned_Vendor),
        Status: getString(item.Status) || 'Draft',
        Budget_Estimate: getNumber(item.Budget_Estimate),
        Category: getString(item.Category),
        Justification: getString(item.Justification),
        Approved_By: getPersonTitle(item.Approved_By),
        Approval_Date: getOptionalString(item.Approval_Date),
      }));

      setItems(requests);
      setFilteredItems(requests);
    } catch (err) {
      console.error('Error fetching procurement requests:', err);
      setError(err instanceof Error ? err.message : 'Failed to load procurement requests');
    } finally {
      setIsLoading(false);
    }
  }, [spHttpClient, pageContext]);

  React.useEffect(() => {
    fetchRequests().catch(() => undefined);
  }, [fetchRequests]);

  // ─── Filtering ─────────────────────────────────────────────────────────────

  React.useEffect(() => {
    let result = items;

    if (statusFilter !== 'All') {
      result = result.filter((item) => item.Status === statusFilter);
    }

    if (searchText) {
      const lower = searchText.toLowerCase();
      result = result.filter(
        (item) =>
          item.RequestID.toLowerCase().includes(lower) ||
          item.Title.toLowerCase().includes(lower) ||
          item.Project_Code.toLowerCase().includes(lower) ||
          item.Category.toLowerCase().includes(lower)
      );
    }

    setFilteredItems(result);
  }, [searchText, statusFilter, items]);

  // ─── Command Bar ───────────────────────────────────────────────────────────

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'newRequest',
      text: 'New Procurement Request',
      iconProps: { iconName: 'Add' },
      onClick: () => setIsNewRequestPanelOpen(true),
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: async () => {
        await fetchRequests();
        onRefresh();
      },
    },
  ];

  // ─── Row Click ─────────────────────────────────────────────────────────────

  const handleRowClick = (item: IProcurementRequest): void => {
    setSelectedRequest(item);
    setIsDetailPanelOpen(true);
  };

  // ─── Panel Handlers ────────────────────────────────────────────────────────

  const handleDetailPanelDismiss = (): void => {
    setIsDetailPanelOpen(false);
    setSelectedRequest(undefined);
  };

  const handleNewRequestDismiss = (): void => {
    setIsNewRequestPanelOpen(false);
  };

  const handleRequestCreated = async (): Promise<void> => {
    setIsNewRequestPanelOpen(false);
    await fetchRequests();
  };

  const handleBidUpdated = async (): Promise<void> => {
    handleDetailPanelDismiss();
    await fetchRequests();
  };

  // ─── Render ────────────────────────────────────────────────────────────────

  if (isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ padding: '60px' }}>
        <Spinner size={SpinnerSize.large} label="Loading procurement requests..." />
      </Stack>
    );
  }

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
        {error}
      </MessageBar>
    );
  }

  // Calculate stats
  const activeRequests = items.filter((i) => !['Cancelled', 'Awarded'].includes(i.Status)).length;
  const bidsInProgress = items.filter((i) => i.Status === 'Bid In Progress').length;
  const awaitingAward = items.filter((i) => i.Status === 'Quotes Received').length;
  const completed = items.filter((i) => i.Status === 'Awarded').length;

  return (
    <div className={classNames.container}>
      {/* Stats Row */}
      <div className={classNames.statsRow}>
        <div className={classNames.statCard}>
          <Icon iconName="ReceiptProcessing" style={{ fontSize: 20, color: '#2563EB' }} />
          <div>
            <div className={classNames.statValue}>{activeRequests}</div>
            <div className={classNames.statLabel}>Active Requests</div>
          </div>
        </div>
        <div className={classNames.statCard}>
          <Icon iconName="Send" style={{ fontSize: 20, color: '#F59E0B' }} />
          <div>
            <div className={classNames.statValue}>{bidsInProgress}</div>
            <div className={classNames.statLabel}>Bids In Progress</div>
          </div>
        </div>
        <div className={classNames.statCard}>
          <Icon iconName="CompareUneven" style={{ fontSize: 20, color: '#8B5CF6' }} />
          <div>
            <div className={classNames.statValue}>{awaitingAward}</div>
            <div className={classNames.statLabel}>Awaiting Award</div>
          </div>
        </div>
        <div className={classNames.statCard}>
          <Icon iconName="Trophy2Solid" style={{ fontSize: 20, color: '#10B981' }} />
          <div>
            <div className={classNames.statValue}>{completed}</div>
            <div className={classNames.statLabel}>Awarded</div>
          </div>
        </div>
      </div>

      {/* Filter + Search Row */}
      <div className={classNames.headerRow}>
        <div className={classNames.filterRow}>
          <SearchBox
            placeholder="Search requests..."
            onChange={(_, newValue) => setSearchText(newValue || '')}
            styles={{ root: { width: isMobileView ? '100%' : '280px' } }}
          />
          <Dropdown
            placeholder="Filter by status"
            selectedKey={statusFilter}
            options={STATUS_FILTER_OPTIONS}
            onChange={(_, option) => setStatusFilter(option?.key as string || 'All')}
            styles={{ root: { width: '160px' } }}
          />
        </div>
        <CommandBar items={commandBarItems} styles={{ root: { padding: 0 } }} />
      </div>

      {/* Data Grid using EnhancedDataGrid */}
      {filteredItems.length === 0 ? (
        <div className={classNames.emptyState}>
          <Icon iconName="ReceiptProcessing" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No procurement requests found</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4 }}>
            {searchText || statusFilter !== 'All'
              ? 'Try adjusting your search or filter criteria'
              : 'Create a new procurement request to get started'}
          </Text>
        </div>
      ) : (
        <div className={classNames.gridContainer}>
          <EnhancedDataGrid
            listName="Procurement_Requests"
            columns={[
              { key: 'RequestID', name: 'Request ID', fieldName: 'RequestID', minWidth: 110, onRender: (item) => <span style={{ fontWeight: 600, fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontSize: 13 }}>{item.RequestID}</span> },
              { key: 'Title', name: 'Description', fieldName: 'Title', minWidth: 200 },
              { key: 'Project_Code', name: 'Project', fieldName: 'Project_Code', minWidth: 100 },
              { key: 'Sourcing_Method', name: 'Sourcing', fieldName: 'Sourcing_Method', minWidth: 120, onRender: (item) => <Tag text={item.Sourcing_Method} /> },
              { key: 'Category', name: 'Category', fieldName: 'Category', minWidth: 90 },
              { key: 'Budget_Estimate', name: 'Budget', fieldName: 'Budget_Estimate', minWidth: 100, onRender: (item) => formatCurrency(item.Budget_Estimate) },
              { key: 'Required_Date', name: 'Required By', fieldName: 'Required_Date', minWidth: 100, onRender: (item) => formatDate(item.Required_Date) },
              { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 120, onRender: (item) => <Tag text={item.Status} /> },
            ]}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            showExport={false}
            showSearch={false}
            showRefresh={false}
            showNewRecord={false}
            onRowDoubleClick={handleRowClick}
          />
        </div>
      )}

      {/* Bid Detail Panel */}
      <BidDetailPanel
        isOpen={isDetailPanelOpen}
        request={selectedRequest}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        onDismiss={handleDetailPanelDismiss}
        onBidUpdated={handleBidUpdated}
      />

      {/* New Procurement Request Panel */}
      <NewProcurementRequestPanel
        isOpen={isNewRequestPanelOpen}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        onDismiss={handleNewRequestDismiss}
        onRequestCreated={handleRequestCreated}
      />
    </div>
  );
};

export default BidManagementTab;
