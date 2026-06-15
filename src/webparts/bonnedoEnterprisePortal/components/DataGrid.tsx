import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode,
  Stack,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  mergeStyleSets,
  getTheme,
  Text,
  Icon,
  TextField,
  IconButton,
  PrimaryButton,
  DefaultButton,
} from '@fluentui/react';
import type { IDetailsListProps } from '@fluentui/react';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';

export interface IDataGridColumn extends IColumn {
  fieldName: string;
  name: string;
  minWidth: number;
  maxWidth?: number;
  isResizable?: boolean;
  isSorted?: boolean;
  isSortedDescending?: boolean;
}

export interface IDataGridProps {
  listName: string;
  columns: IDataGridColumn[];
  filterQuery?: string;
  expandQuery?: string;
  pageSize?: number;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  onRowSelected?: (record: IListItem) => void;
  onRowDoubleClick?: (record: IListItem) => void;
  showMobileCards?: boolean;
  // Toolbar props
  title?: string;
  showBackButton?: boolean;
  showSearch?: boolean;
  showRefresh?: boolean;
  showNewRecord?: boolean;
  onBack?: () => void;
  onRefresh?: () => void;
  onNewRecord?: () => void;
}

interface IDataGridState {
  items: IListItem[];
  filteredItems: IListItem[];
  columns: IDataGridColumn[];
  isLoading: boolean;
  error: string | undefined;
  currentPage: number;
  pageSize: number;
  totalItems: number;
  sortedColumn: string | undefined;
  sortDescending: boolean;
  searchQuery: string;
  isMobile: boolean;
}

const DataGrid: React.FC<IDataGridProps> = ({
  listName,
  columns,
  filterQuery,
  expandQuery,
  pageSize = 10,
  spHttpClient,
  pageContext,
  onRowSelected,
  onRowDoubleClick,
  showMobileCards = true,
  // Toolbar props
  title,
  showBackButton = false,
  showSearch = true,
  showRefresh = true,
  showNewRecord = false,
  onBack,
  onRefresh,
  onNewRecord,
}) => {
  const [state, setState] = React.useState<IDataGridState>({
    items: [],
    filteredItems: [],
    columns: columns,
    isLoading: true,
    error: undefined,
    currentPage: 1,
    pageSize: pageSize,
    totalItems: 0,
    sortedColumn: undefined,
    sortDescending: false,
    searchQuery: '',
    isMobile: false,
  });

  // Check for mobile viewport
  React.useEffect(() => {
    const checkMobile = (): void => {
      const width = window.innerWidth;
      setState((prev) => ({ ...prev, isMobile: width < 768 }));
    };

    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  // Handle search/filter
  React.useEffect(() => {
    if (!state.searchQuery.trim()) {
      setState((prev) => ({
        ...prev,
        filteredItems: state.items,
      }));
      return;
    }

    const searchLower = state.searchQuery.toLowerCase();
    const filtered = state.items.filter((item) => {
      return Object.keys(item).some((key) => {
        const value = item[key];
        if (value === null || value === undefined) return false;
        if (typeof value === 'object') return false;
        return String(value).toLowerCase().includes(searchLower);
      });
    });

    setState((prev) => ({
      ...prev,
      filteredItems: filtered,
    }));
  }, [state.searchQuery, state.items]);

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const selection = React.useRef(
    new Selection({
      onSelectionChanged: () => {
        const selectedItems = selection.current.getSelection();
        if (selectedItems.length > 0 && onRowSelected) {
          onRowSelected(selectedItems[0] as IListItem);
        }
      },
    })
  );

  // Fetch data from SharePoint
  const fetchData = React.useCallback(async (page: number = 1) => {
    try {
      setState((prev) => ({ ...prev, isLoading: true, error: undefined }));

      const skip = (page - 1) * state.pageSize;
      const items = await sharePointService.getListData(
        listName,
        filterQuery,
        state.pageSize,
        skip,
        expandQuery
      );

      const totalItems = await sharePointService.getListItemCount(listName, filterQuery);

      console.log('DataGrid items received:', items);

      setState((prev) => ({
        ...prev,
        items,
        filteredItems: items,
        totalItems,
        currentPage: page,
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
  }, [listName, filterQuery, expandQuery, sharePointService, state.pageSize]);

  // Initial data load
  React.useEffect(() => {
    fetchData(1).then((data) => {
      console.log('DataGrid fetched data:', data);
    }).catch((err) => {
      console.error('Error fetching initial data:', err);
    });
  }, [listName, filterQuery, expandQuery]);

  // Handle column header click for sorting
  const handleColumnClick = (column: IDataGridColumn): void => {
    const isSortedDescending = !column.isSortedDescending;
    const newColumns = state.columns.map((col) => ({
      ...col,
      isSorted: col.key === column.key,
      isSortedDescending:
        col.key === column.key ? isSortedDescending : false,
    }));

    const sortedItems = [...state.filteredItems].sort((a, b) => {
      const aValue = a[column.fieldName];
      const bValue = b[column.fieldName];

      if (aValue === bValue) return 0;

      const comparison = aValue < bValue ? -1 : 1;
      return isSortedDescending ? -comparison : comparison;
    });

    setState((prev) => ({
      ...prev,
      columns: newColumns,
      filteredItems: sortedItems,
      sortedColumn: column.key,
      sortDescending: isSortedDescending,
    }));
  };

  // Handle pagination
  const handlePageChange = (pageNumber: number): void => {
    fetchData(pageNumber).catch((err) => {
      console.error('Error fetching data:', err);
    });
    selection.current.setAllSelected(false);
  };

  // Handle row double click
  const handleRowDoubleClick = (item: IListItem): void => {
    if (onRowDoubleClick) {
      onRowDoubleClick(item);
    }
  };

  // Handle card click
  const handleCardClick = (item: IListItem): void => {
    if (onRowDoubleClick) {
      onRowDoubleClick(item);
    }
  };

  // Get status color
  const getStatusColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'approved':
        return '#107c10';
      case 'rejected':
        return '#a80000';
      case 'pending':
        return '#ca5010';
      case 'draft':
        return '#605e5c';
      case 'active':
        return '#107c10';
      case 'completed':
        return '#0078d4';
      default:
        return '#605e5c';
    }
  };

  const getStatusBackgroundColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'approved':
        return '#dff6dd';
      case 'rejected':
        return '#fde7e9';
      case 'pending':
        return '#fff4ce';
      case 'draft':
        return '#f3f2f1';
      case 'active':
        return '#dff6dd';
      case 'completed':
        return '#e6f2ff';
      default:
        return '#f3f2f1';
    }
  };

  const theme = getTheme();

  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      backgroundColor: theme.palette.white,
    },
    // Toolbar styles
    toolbar: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      padding: '12px 16px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      flexWrap: 'wrap',
      gap: '12px',
    },
    toolbarLeft: {
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
    },
    toolbarRight: {
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
    },
    toolbarTitle: {
      fontWeight: 600,
      fontSize: '18px',
      color: theme.palette.neutralPrimary,
    },
    searchField: {
      width: '250px',
    },
    container: {
      flex: 1,
      overflow: 'hidden',
      display: 'flex',
      flexDirection: 'column',
    },
    tableContainer: {
      flex: 1,
      overflow: 'auto',
      minHeight: 0, // Allow flex shrinking
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: '2px',
    },
    detailsList: {
      flex: 1,
      overflow: 'auto',
    },
    paginationContainer: {
      padding: '16px',
      borderTop: `1px solid ${theme.palette.neutralLight}`,
      backgroundColor: theme.palette.neutralLighterAlt,
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
    },
    errorContainer: {
      padding: '16px',
    },
    loadingContainer: {
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      height: '300px',
    },
    emptyContainer: {
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      alignItems: 'center',
      height: '300px',
      width: '100%',
      color: theme.palette.neutralSecondary,
      backgroundColor: theme.palette.neutralLighterAlt,
      border: `1px dashed ${theme.palette.neutralLight}`,
      borderRadius: '4px',
      margin: '16px',
    },
    // Mobile card styles
    cardContainer: {
      display: 'flex',
      flexDirection: 'column',
      gap: '12px',
      padding: '12px',
      overflowY: 'auto',
      minHeight: 0,
    },
    card: {
      backgroundColor: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: '8px',
      padding: '16px',
      cursor: 'pointer',
      boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
    },
    cardHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '8px',
    },
    cardTitle: {
      fontWeight: 600,
      fontSize: '16px',
      color: theme.palette.neutralPrimary,
    },
    cardStatus: {
      display: 'inline-flex',
      alignItems: 'center',
      padding: '2px 8px',
      borderRadius: '10px',
      fontSize: '11px',
      fontWeight: 600,
      textTransform: 'uppercase',
    },
    cardFields: {
      display: 'flex',
      flexDirection: 'column',
      gap: '4px',
    },
    cardField: {
      display: 'flex',
      justifyContent: 'space-between',
      fontSize: '13px',
    },
    cardFieldLabel: {
      color: theme.palette.neutralSecondary,
    },
    cardFieldValue: {
      color: theme.palette.neutralPrimary,
      fontWeight: 500,
    },
  });

  const totalPages = Math.ceil(state.totalItems / state.pageSize);

  const renderLookupValue = (item: IListItem, fieldName: string): React.ReactNode => {
    const value = item[fieldName];
    if (value && typeof value === 'object') {
      // Support both standard SP Title/Name and app-specific master list fields (Material_Name, Project_Code, Vendor_Name, etc.)
      return value.Title || value.Name || value.Material_Name || value.Project_Code || value.Vendor_Name || value.Vendor || value.Code || value.ID || '';
    }
    return value ?? '';
  };

  // Adjust columns for mobile responsiveness - ensure minimum readable width
  const displayColumns = state.isMobile
    ? state.columns.map(col => ({
        ...col,
        minWidth: Math.max(col.minWidth * 0.6, 70),
        maxWidth: Math.min(col.maxWidth || col.minWidth * 2, window.innerWidth * 0.25),
        onRender: col.onRender || ((item: IListItem) => renderLookupValue(item, col.fieldName)),
      }))
    : state.columns.map(col => ({
        ...col,
        onRender: col.onRender || ((item: IListItem) => renderLookupValue(item, col.fieldName)),
      }));

  // Render mobile card
  const renderCard = (item: IListItem): React.ReactNode => {
    const statusField = item.Status || item.Approval_Status || item.status;
    const titleField = item.Title || item.Project_Code || item.Name || 'Record';

    // Get first 3-4 key fields to display
    const displayFields = state.columns
      .filter(col =>
        col.fieldName !== 'ID' &&
        col.fieldName !== 'Title' &&
        col.fieldName !== 'Status' &&
        col.fieldName !== 'Approval_Status' &&
        col.fieldName !== 'status'
      )
      .slice(0, 4);

    return (
      <div
        key={item.ID}
        className={classNames.card}
        onClick={() => handleCardClick(item)}
      >
        <div className={classNames.cardHeader}>
          <Text className={classNames.cardTitle}>{titleField}</Text>
          {statusField && (
            <span
              className={classNames.cardStatus}
              style={{
                backgroundColor: getStatusBackgroundColor(statusField),
                color: getStatusColor(statusField),
              }}
            >
              {statusField}
            </span>
          )}
        </div>
        <div className={classNames.cardFields}>
          {displayFields.map((col) => {
            const value = item[col.fieldName];
            if (!value) return null;
            return (
              <div key={col.fieldName} className={classNames.cardField}>
                <Text className={classNames.cardFieldLabel}>{col.name}</Text>
                <Text className={classNames.cardFieldValue}>
                  {typeof value === 'object' ? JSON.stringify(value) : String(value)}
                </Text>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  return (
    <div className={`${classNames.root} dataGridHost`}>
      {state.error && (
        <div className={classNames.errorContainer}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {state.error}
          </MessageBar>
        </div>
      )}

      {/* Toolbar */}
      {(title || showBackButton || showSearch || showRefresh || showNewRecord) && (
        <div className={classNames.toolbar}>
          <div className={classNames.toolbarLeft}>
            {showBackButton && (
              <IconButton
                iconProps={{ iconName: 'Back' }}
                title="Back"
                onClick={onBack}
              />
            )}
            {title && (
              <Text className={classNames.toolbarTitle}>{title}</Text>
            )}
          </div>
          <div className={classNames.toolbarRight}>
            {showSearch && (
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <TextField
                  className={classNames.searchField}
                  placeholder="Search..."
                  value={state.searchQuery}
                  onChange={(_, value) => setState((prev) => ({ ...prev, searchQuery: value || '' }))}
                  iconProps={{ iconName: 'Search' }}
                />
                {state.searchQuery && (
                  <IconButton
                    iconProps={{ iconName: 'Cancel' }}
                    title="Clear search"
                    onClick={() => setState((prev) => ({ ...prev, searchQuery: '' }))}
                  />
                )}
              </Stack>
            )}
            {showRefresh && (
              <IconButton
                iconProps={{ iconName: 'Refresh' }}
                title="Refresh"
                onClick={() => {
                  if (onRefresh) {
                    onRefresh();
                  } else {
                    fetchData(1);
                  }
                }}
              />
            )}
            {showNewRecord && (
              <PrimaryButton
                text="New Record"
                iconProps={{ iconName: 'Add' }}
                onClick={onNewRecord}
              />
            )}
          </div>
        </div>
      )}

      <div className={classNames.container}>
        {state.isLoading ? (
          <div className={classNames.loadingContainer}>
            <Spinner size={SpinnerSize.large} label="Loading data..." />
          </div>
        ) : state.filteredItems.length === 0 ? (
          <div className={classNames.emptyContainer}>
            <Icon iconName="ClipboardEmpty" style={{ fontSize: '48px', marginBottom: '16px' }} />
            <Text variant="large" style={{ marginBottom: '8px' }}>No items found</Text>
            <Text variant="medium" style={{ color: theme.palette.neutralSecondary }}>
              Click "+ New Record" to add your first item
            </Text>
          </div>
        ) : (
          <>
            {/* Responsive Table Layout - shows table on all screen sizes with horizontal scroll on mobile */}
            <div
              className={classNames.tableContainer}
              style={{
                overflowX: state.isMobile ? 'auto' : 'visible',
                overflowY: 'auto',
                minWidth: state.isMobile ? '100%' : 'auto',
              }}
            >
              <div
                className="dataGridTableWrapper"
                style={{
                  minWidth: state.isMobile ? `${displayColumns.reduce((sum, col) => sum + col.minWidth, 0)}px` : 'auto',
                  isolation: 'isolate' // CSS isolation to prevent external style interference
                }}
              >
                <DetailsList
                  items={state.filteredItems}
                  columns={displayColumns}
                  selectionMode={SelectionMode.single}
                  selection={selection.current}
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  isHeaderVisible={true}
                  onColumnHeaderClick={(_, column) => {
                    if (column) {
                      handleColumnClick(column as IDataGridColumn);
                    }
                  }}
                  onItemInvoked={handleRowDoubleClick}
                />
              </div>
            </div>

            {totalPages > 1 && (
              <div className={classNames.paginationContainer}>
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  <Text variant="small">
                    Page {state.currentPage} of {totalPages} | Total: {state.totalItems} items
                  </Text>
                </Stack>
                <Pagination
                  currentPage={state.currentPage}
                  totalPages={totalPages}
                  onChange={handlePageChange}
                />
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};

export default DataGrid;
