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
import VendorDetailPanel from './VendorDetailPanel';
import VendorRegistrationPanel from './VendorRegistrationPanel';
import EnhancedDataGrid from '../EnhancedDataGrid';
import { IDataGridColumn } from '../EnhancedDataGrid';
import { Tag, formatDate } from '../TagRenderer';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IVendorRegistrationTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
}

export interface IVendorItem {
  ID: number;
  Vendor_Code: string;
  Vendor_Name: string;
  Contact_Person: string;
  Email: string;
  Phone: string;
  Address: string;
  Registration_Type: string;
  Vendor_Category: string;
  Tax_ID: string;
  RC_Number: string;
  Bank_Name: string;
  Bank_Account: string;
  Vendor_Status: string;
  Approval_Date: string | undefined;
  Approved_By: string | undefined;
  Notes: string;
  Created: string;
  Modified: string;
}

// ─── Status Color Map ────────────────────────────────────────────────────────

const VENDOR_STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  Active: { bg: '#D1FAE5', text: '#065F46' },
  Inactive: { bg: '#F3F4F6', text: '#6B7280' },
  Pending: { bg: '#FEF3C7', text: '#92400E' },
  'Under Review': { bg: '#DBEAFE', text: '#1E40AF' },
  Approved: { bg: '#D1FAE5', text: '#065F46' },
  Rejected: { bg: '#FEE2E2', text: '#991B1B' },
  Suspended: { bg: '#FFEDD5', text: '#9A3412' },
  Blacklisted: { bg: '#FCE7F3', text: '#831843' },
};

const STATUS_FILTER_OPTIONS: IDropdownOption[] = [
  { key: 'All', text: 'All Statuses' },
  { key: 'Pending', text: 'Pending' },
  { key: 'Under Review', text: 'Under Review' },
  { key: 'Active', text: 'Active' },
  { key: 'Inactive', text: 'Inactive' },
  { key: 'Rejected', text: 'Rejected' },
  { key: 'Suspended', text: 'Suspended' },
  { key: 'Blacklisted', text: 'Blacklisted' },
];

// ─── Styles ──────────────────────────────────────────────────────────────────

interface ISharePointListItem {
  ID: number;
  [key: string]: unknown;
}

const getString = (value: unknown): string => typeof value === 'string' ? value : '';
const getOptionalString = (value: unknown): string | undefined => typeof value === 'string' && value ? value : undefined;

const getPersonTitle = (value: unknown): string | undefined => {
  if (value && typeof value === 'object' && 'Title' in value) {
    const title = (value as { Title?: unknown }).Title;
    return typeof title === 'string' ? title : undefined;
  }

  return undefined;
};

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
    vendorRow: {
      cursor: 'pointer',
      selectors: {
        '&:hover': {
          backgroundColor: '#F8FAFC',
        },
      },
    },
    gridContainer: { flex: 1, display: 'flex', flexDirection: 'column' },
    emptyState: {
      textAlign: 'center',
      padding: '60px 20px',
      color: '#5A6A85',
    },
  });

// ─── Component ───────────────────────────────────────────────────────────────

const VendorRegistrationTab: React.FC<IVendorRegistrationTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
}) => {
  const [items, setItems] = React.useState<IVendorItem[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IVendorItem[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [searchText, setSearchText] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<string>('All');
  const [selectedVendor, setSelectedVendor] = React.useState<IVendorItem | undefined>(undefined);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [isRegistrationPanelOpen, setIsRegistrationPanelOpen] = React.useState(false);

  const classNames = getStyles() as unknown as { [key: string]: string };

  // ─── Data Fetching ─────────────────────────────────────────────────────────

  const fetchVendors = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(undefined);
    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Vendors_Master')/items?$top=500&$orderby=Created desc`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch vendors: ${response.status}`);
      }

      const data = await response.json();
      const vendors: IVendorItem[] = (data.value || []).map((item: ISharePointListItem) => ({
        ID: item.ID,
        Vendor_Code: getString(item.Vendor_Code),
        Vendor_Name: getString(item.Vendor_Name),
        Contact_Person: getString(item.Contact_Person),
        Email: getString(item.Email),
        Phone: getString(item.Phone),
        Address: getString(item.Address),
        Registration_Type: getString(item.Registration_Type),
        Vendor_Category: getString(item.Vendor_Category),
        Tax_ID: getString(item.Tax_ID),
        RC_Number: getString(item.RC_Number),
        Bank_Name: getString(item.Bank_Name),
        Bank_Account: getString(item.Bank_Account),
        Vendor_Status: getString(item.Vendor_Status) || 'Pending',
        Approval_Date: getOptionalString(item.Approval_Date),
        Approved_By: getPersonTitle(item.Approved_By),
        Notes: getString(item.Notes),
        Created: getString(item.Created),
        Modified: getString(item.Modified),
      }));

      setItems(vendors);
      setFilteredItems(vendors);
    } catch (err) {
      console.error('Error fetching vendors:', err);
      setError(err instanceof Error ? err.message : 'Failed to load vendors');
    } finally {
      setIsLoading(false);
    }
  }, [spHttpClient, pageContext]);

  React.useEffect(() => {
    fetchVendors().catch(() => undefined);
  }, [fetchVendors]);

  // ─── Filtering ─────────────────────────────────────────────────────────────

  React.useEffect(() => {
    let result = items;

    // Status filter
    if (statusFilter !== 'All') {
      result = result.filter((item) => item.Vendor_Status === statusFilter);
    }

    // Search filter
    if (searchText) {
      const lower = searchText.toLowerCase();
      result = result.filter(
        (item) =>
          item.Vendor_Code.toLowerCase().includes(lower) ||
          item.Vendor_Name.toLowerCase().includes(lower) ||
          item.Contact_Person.toLowerCase().includes(lower) ||
          item.Email.toLowerCase().includes(lower) ||
          item.Vendor_Category.toLowerCase().includes(lower)
      );
    }

    setFilteredItems(result);
  }, [searchText, statusFilter, items]);

  // ─── Command Bar ───────────────────────────────────────────────────────────

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'registerVendor',
      text: 'Register Vendor',
      iconProps: { iconName: 'AddFriend' },
      onClick: () => {
        setIsRegistrationPanelOpen(true);
      },
    },
    {
      key: 'sendLink',
      text: 'Send Registration Link',
      iconProps: { iconName: 'Share' },
      onClick: async () => {
        const formLink = `${pageContext.web.absoluteUrl}/Lists/ENT_Vendors_Master/NewForm.aspx?Source=${encodeURIComponent(window.location.href)}`;
        try {
          await navigator.clipboard.writeText(formLink);
          alert('Registration link copied to clipboard. Send to vendor via email.');
        } catch {
          window.open(`mailto:?subject=Bonnedo Vendor Registration&body=Please register as a vendor using this link: ${formLink}`, '_blank');
        }
      },
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: async () => {
        await fetchVendors();
        onRefresh();
      },
    },
  ];

  // ─── Row Click Handler ─────────────────────────────────────────────────────

  const handleRowClick = (item: IVendorItem): void => {
    setSelectedVendor(item);
    setIsPanelOpen(true);
  };

  // ─── Panel Handlers ────────────────────────────────────────────────────────

  const handlePanelDismiss = (): void => {
    setIsPanelOpen(false);
    setSelectedVendor(undefined);
  };

  const handleRegistrationPanelDismiss = (): void => {
    setIsRegistrationPanelOpen(false);
  };

  const handleVendorRegistered = (): void => {
    setIsRegistrationPanelOpen(false);
    fetchVendors().catch(() => undefined);
  };

  const handleVendorUpdated = async (): Promise<void> => {
    handlePanelDismiss();
    await fetchVendors();
  };

  // ─── Render ────────────────────────────────────────────────────────────────

  if (isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ padding: '60px' }}>
        <Spinner size={SpinnerSize.large} label="Loading vendors..." />
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
  const totalVendors = items.length;
  const approvedCount = items.filter((i) => i.Vendor_Status === 'Active').length;
  const pendingCount = items.filter((i) => i.Vendor_Status === 'Pending' || i.Vendor_Status === 'Under Review').length;
  const rejectedCount = items.filter((i) => i.Vendor_Status === 'Rejected' || i.Vendor_Status === 'Blacklisted').length;

  return (
    <div className={classNames.container}>
      {/* Stats Row */}
      <div className={classNames.statsRow}>
        <div className={classNames.statCard}>
          <Icon iconName="ContactList" style={{ fontSize: 20, color: '#2563EB' }} />
          <div>
            <div className={classNames.statValue}>{totalVendors}</div>
            <div className={classNames.statLabel}>Total Vendors</div>
          </div>
        </div>
        <div className={classNames.statCard}>
          <Icon iconName="CompletedSolid" style={{ fontSize: 20, color: '#10B981' }} />
          <div>
            <div className={classNames.statValue}>{approvedCount}</div>
            <div className={classNames.statLabel}>Approved</div>
          </div>
        </div>
        <div className={classNames.statCard}>
          <Icon iconName="Clock" style={{ fontSize: 20, color: '#F59E0B' }} />
          <div>
            <div className={classNames.statValue}>{pendingCount}</div>
            <div className={classNames.statLabel}>Pending</div>
          </div>
        </div>
        <div className={classNames.statCard}>
          <Icon iconName="ErrorBadge" style={{ fontSize: 20, color: '#EF4444' }} />
          <div>
            <div className={classNames.statValue}>{rejectedCount}</div>
            <div className={classNames.statLabel}>Rejected</div>
          </div>
        </div>
      </div>

      {/* Filter + Search Row */}
      <div className={classNames.headerRow}>
        <div className={classNames.filterRow}>
          <SearchBox
            placeholder="Search vendors..."
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
          <Icon iconName="ContactList" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No vendors found</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4 }}>
            {searchText || statusFilter !== 'All'
              ? 'Try adjusting your search or filter criteria'
              : 'Register vendors to get started'}
          </Text>
        </div>
      ) : (
        <div className={classNames.gridContainer}>
          <EnhancedDataGrid
            listName="ENT_Vendors_Master"
            columns={[
              { key: 'Vendor_Code', name: 'Vendor ID', fieldName: 'Vendor_Code', minWidth: 110, onRender: (item) => <span style={{ fontWeight: 600, fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontSize: 13 }}>{item.Vendor_Code}</span> },
              { key: 'Vendor_Name', name: 'Company Name', fieldName: 'Vendor_Name', minWidth: 200 },
              { key: 'Contact_Person', name: 'Contact', fieldName: 'Contact_Person', minWidth: 130 },
              { key: 'Vendor_Category', name: 'Category', fieldName: 'Vendor_Category', minWidth: 120, onRender: (item) => <Tag text={item.Vendor_Category} /> },
              { key: 'Registration_Type', name: 'Type', fieldName: 'Registration_Type', minWidth: 90 },
              { key: 'Vendor_Status', name: 'Status', fieldName: 'Vendor_Status', minWidth: 110, onRender: (item) => <Tag text={item.Vendor_Status} /> },
              { key: 'Created', name: 'Registered', fieldName: 'Created', minWidth: 100, onRender: (item) => formatDate(item.Created) },
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

      {/* Vendor Detail Panel */}
      <VendorRegistrationPanel
        isOpen={isRegistrationPanelOpen}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        onDismiss={handleRegistrationPanelDismiss}
        onVendorRegistered={handleVendorRegistered}
      />

      <VendorDetailPanel
        isOpen={isPanelOpen}
        vendor={selectedVendor}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        onDismiss={handlePanelDismiss}
        onVendorUpdated={handleVendorUpdated}
      />
    </div>
  );
};

export default VendorRegistrationTab;
