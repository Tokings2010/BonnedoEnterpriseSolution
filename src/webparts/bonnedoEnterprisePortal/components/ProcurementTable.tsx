import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode,
  TextField,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Text,
  getTheme,
  mergeStyleSets,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../services/SharePointService';
//// Removed unused styles import

export interface IProcurementTableProps {
  listName: string;
  recordType: 'MR' | 'PR' | 'PO' | 'GRN';
  columns: IColumn[];
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  onRowSelected?: (record: any) => void;
  filterQuery?: string;
  pageSize?: number;
}

interface IProcurementTableState {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  items: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  filteredItems: any[];
  isLoading: boolean;
  error: string | undefined;
  searchText: string;
  statusFilter: string;
  isMobile: boolean;
}

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'All Statuses' },
  { key: 'Draft', text: 'Draft' },
  { key: 'Submitted', text: 'Submitted' },
  { key: 'Approved', text: 'Approved' },
  { key: 'Rejected', text: 'Rejected' },
  { key: 'Completed', text: 'Completed' },
];

const ProcurementTable: React.FC<IProcurementTableProps> = ({
  listName,
  recordType,
  columns,
  spHttpClient,
  pageContext,
  onRowSelected,
  filterQuery,
  pageSize = 20,
}) => {
  const [state, setState] = React.useState<IProcurementTableState>({
    items: [],
    filteredItems: [],
    isLoading: true,
    error: undefined,
    searchText: '',
    statusFilter: '',
    isMobile: window.innerWidth < 768,
  });

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const selection = React.useRef(
    new Selection({
      onSelectionChanged: () => {
        const selectedItems = selection.current.getSelection();
        if (selectedItems.length > 0 && onRowSelected) {
          onRowSelected(selectedItems[0]);
        }
      },
    })
  );

  // Fetch data
  React.useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        setState((prev) => ({ ...prev, isLoading: true, error: undefined }));

        let query = filterQuery || '';
        if (state.statusFilter) {
          const statusCondition = `Status eq '${state.statusFilter}'`;
          query = query ? `${query} and ${statusCondition}` : statusCondition;
        }

        const items = await sharePointService.getListData(listName, query, pageSize);
        setState((prev) => ({
          ...prev,
          items,
          filteredItems: items,
          isLoading: false,
        }));
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Failed to load data';
        setState((prev) => ({
          ...prev,
          isLoading: false,
          error: errorMessage,
        }));
      }
    };

    fetchData().catch((err) => {
      console.error('Error fetching procurement data:', err);
    });
  }, [listName, filterQuery, state.statusFilter, sharePointService, pageSize]);

  // Handle search
  React.useEffect(() => {
    const filtered = state.items.filter((item) => {
      const searchLower = state.searchText.toLowerCase();
      return (
        item.Title?.toLowerCase().includes(searchLower) ||
        item.Project_Code?.toLowerCase().includes(searchLower) ||
        item.Material?.toLowerCase().includes(searchLower) ||
        item.Vendor?.toLowerCase().includes(searchLower) ||
        item.Description?.toLowerCase().includes(searchLower)
      );
    });

    setState((prev) => ({
      ...prev,
      filteredItems: filtered,
    }));
  }, [state.searchText, state.items]);

  // Handle window resize
  React.useEffect(() => {
    const handleResize = (): void => {
      setState((prev) => ({
        ...prev,
        isMobile: window.innerWidth < 768,
      }));
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const theme = getTheme();
  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
    },
    filterContainer: {
      padding: '16px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      display: 'flex',
      gap: '12px',
      flexWrap: 'wrap',
      alignItems: 'flex-end',
    },
    tableContainer: {
      flex: 1,
      overflow: 'hidden',
    },
    cardContainer: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
      gap: '16px',
      padding: '16px',
      overflow: 'auto',
    },
    card: {
      padding: '16px',
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: '4px',
      cursor: 'pointer',
      transition: 'all 0.2s ease',
      backgroundColor: theme.palette.white,
      selectors: {
        '&:hover': {
          boxShadow: theme.effects.elevation8,
          transform: 'translateY(-2px)',
        },
      },
    },
    cardField: {
      marginBottom: '12px',
      paddingBottom: '12px',
      borderBottom: `1px solid ${theme.palette.neutralLighter}`,
      selectors: {
        '&:last-child': {
          borderBottom: 'none',
          marginBottom: 0,
          paddingBottom: 0,
        },
      },
    },
    cardLabel: {
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
      fontSize: '12px',
      marginBottom: '4px',
    },
    cardValue: {
      color: theme.palette.neutralSecondary,
      fontSize: '14px',
    },
  });

  if (state.isLoading) {
    return (
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '300px' }}>
        <Spinner size={SpinnerSize.large} label="Loading records..." />
      </div>
    );
  }

  if (state.error) {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline>
        {state.error}
      </MessageBar>
    );
  }

  if (state.filteredItems.length === 0) {
    return (
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '300px' }}>
        <Text variant="medium">No records found</Text>
      </div>
    );
  }

  return (
    <div className={classNames.root}>
      {/* Filter Bar */}
      <div className={classNames.filterContainer}>
        <TextField
          placeholder="Search records..."
          value={state.searchText}
          onChange={(_, value) => setState((prev) => ({ ...prev, searchText: value || '' }))}
          style={{ flex: 1, minWidth: '200px' }}
          iconProps={{ iconName: 'Search' }}
        />
        <Dropdown
          label="Status"
          options={STATUS_OPTIONS}
          selectedKey={state.statusFilter}
          onChange={(_, option) => setState((prev) => ({ ...prev, statusFilter: option?.key as string || '' }))}
          style={{ minWidth: '150px' }}
        />
      </div>

      {/* Table/Card View */}
      {state.isMobile ? (
        // Card Layout for Mobile
        <div className={classNames.cardContainer}>
          {state.filteredItems.map((item) => (
            <div
              key={item.ID}
              className={classNames.card}
              onClick={() => onRowSelected?.(item)}
            >
              <div className={classNames.cardField}>
                <div className={classNames.cardLabel}>Record Number</div>
                <div className={classNames.cardValue}>{item.Title}</div>
              </div>

              {item.Project_Code && (
                <div className={classNames.cardField}>
                  <div className={classNames.cardLabel}>Project Code</div>
                  <div className={classNames.cardValue}>{item.Project_Code}</div>
                </div>
              )}

              {item.Material && (
                <div className={classNames.cardField}>
                  <div className={classNames.cardLabel}>Material</div>
                  <div className={classNames.cardValue}>{item.Material}</div>
                </div>
              )}

              {item.Vendor && (
                <div className={classNames.cardField}>
                  <div className={classNames.cardLabel}>Vendor</div>
                  <div className={classNames.cardValue}>{item.Vendor}</div>
                </div>
              )}

              {item.Quantity && (
                <div className={classNames.cardField}>
                  <div className={classNames.cardLabel}>Quantity</div>
                  <div className={classNames.cardValue}>{item.Quantity} {item.UOM}</div>
                </div>
              )}

              {item.Status && (
                <div className={classNames.cardField}>
                  <div className={classNames.cardLabel}>Status</div>
                  <div className={classNames.cardValue} style={{ fontWeight: 600 }}>
                    {item.Status}
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>
      ) : (
        // DetailsList for Desktop
        <div className={classNames.tableContainer}>
          <DetailsList
            items={state.filteredItems}
            columns={columns}
            selectionMode={SelectionMode.single}
            selection={selection.current}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            isHeaderVisible={true}
            onItemInvoked={(item) => onRowSelected?.(item)}
            styles={{
              root: {
                selectors: {
                  '.ms-DetailsHeader': {
                    borderBottom: `1px solid ${theme.palette.neutralLight}`,
                    padding: '12px 16px 16px 16px',
                    backgroundColor: theme.palette.neutralLighterAlt,
                  },
                  '.ms-DetailsHeader-cell': {
                    padding: '0',
                    fontWeight: 600,
                    color: theme.palette.neutralPrimary,
                    textAlign: 'left',
                  },
                  '.ms-DetailsRow': {
                    borderBottom: `1px solid ${theme.palette.neutralLight}`,
                    marginTop: '4px',
                    selectors: {
                      '&:hover': {
                        backgroundColor: theme.palette.neutralLighterAlt,
                        cursor: 'pointer',
                      },
                    },
                  },
                  '.ms-DetailsRow.is-selected': {
                    backgroundColor: theme.palette.themeLighter,
                  },
                },
              },
            }}
          />
        </div>
      )}
    </div>
  );
};

export default ProcurementTable;
