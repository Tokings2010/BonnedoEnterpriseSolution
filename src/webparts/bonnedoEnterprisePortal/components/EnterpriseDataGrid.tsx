import * as React from 'react';
import {
  DetailsList, DetailsListLayoutMode, Selection, IColumn, SelectionMode,
  Spinner, SpinnerSize, MessageBar, MessageBarType, mergeStyleSets,
  Text, Icon, SearchBox, IconButton, PrimaryButton, DefaultButton,
  Dropdown, IDropdownOption, DatePicker, IDatePickerStyles,
  CheckboxVisibility, ConstrainMode, ScrollablePane, ScrollbarVisibility,
  Sticky, StickyPositionType, TooltipHost, ITooltipHostStyles,
  DirectionalHint, ContextualMenu, IContextualMenuProps, IContextualMenuItem,
  ActionButton, CommandBarButton,
  Panel, PanelType,
  Label,
  FontIcon,
  ITag,
} from '@fluentui/react';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';
import { exportToCsv } from './TagRenderer';

// ─── Exported Types (backward-compatible with EnhancedDataGrid) ──────────────

export interface IDataGridColumn extends IColumn {
  fieldName: string;
  name: string;
  minWidth: number;
  maxWidth?: number;
  isResizable?: boolean;
  isSorted?: boolean;
  isSortedDescending?: boolean;
  /** Custom render for the column */
  onRender?: (item: IListItem) => React.ReactNode;
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
  title?: string;
  showBackButton?: boolean;
  showSearch?: boolean;
  showRefresh?: boolean;
  showNewRecord?: boolean;
  showExport?: boolean;
  onBack?: () => void;
  onRefresh?: () => void;
  onNewRecord?: () => void;
  newRecordLabel?: string;
  customFilters?: React.ReactNode;
  statusFilterOptions?: IDropdownOption[];
  statusFilterValue?: string;
  onStatusFilterChange?: (value: string) => void;
  showDateFilter?: boolean;
  dateFieldLabel?: string;
  dateStart?: Date;
  dateEnd?: Date;
  onDateChange?: (start: Date | undefined, end: Date | undefined) => void;
  className?: string;
  sortOptions?: IDropdownOption[];
  sortField?: string;
  onSortChange?: (field: string) => void;
  /** Show density toggle (compact/comfortable) */
  showDensityToggle?: boolean;
  /** Enable column visibility customization */
  showColumnChooser?: boolean;
  /** Enable per-column filter menus */
  showColumnFilters?: boolean;
}

// ─── Internal Types ──────────────────────────────────────────────────────────

interface IEnterpriseDataGridState {
  items: IListItem[];
  filteredItems: IListItem[];
  columns: IDataGridColumn[];
  isLoading: boolean;
  error: string | undefined;
  currentPage: number;
  pageSize: number;
  totalItems: number;
  searchQuery: string;
  isMobile: boolean;
  isCompact: boolean;
  columnFilterMenus: Record<string, string[]>;
  columnSearchTerms: Record<string, string>;
  visibleColumnKeys: Set<string>;
  isColumnChooserOpen: boolean;
  densityMode: 'comfortable' | 'compact';
}

// ─── Styles ──────────────────────────────────────────────────────────────────

const THEME = {
  primary: '#0078D4',
  primaryDark: '#106EBE',
  primaryLight: 'rgba(0,120,212,0.06)',
  textPrimary: '#242424',
  textSecondary: '#616161',
  textTertiary: '#8A8A8A',
  border: '#E1E1E1',
  borderLight: '#EDEDED',
  background: '#FFFFFF',
  backgroundAlt: '#FAFAFA',
  headerBg: '#FFFFFF',
  headerBorder: '#E1E1E1',
  danger: '#D13438',
  warning: '#F2C811',
  success: '#107C10',
  info: '#0078D4',
  shadow: '0 1px 2px rgba(0,0,0,0.06)',
  fontFamily: "'Segoe UI', -apple-system, 'Helvetica Neue', sans-serif",
  monoFont: "'Cascadia Code', 'Fira Code', Consolas, monospace",
};

const dateStyles: Partial<IDatePickerStyles> = {
  root: { maxWidth: 140, minWidth: 120 },
  textField: { border: `1px solid ${THEME.border}`, borderRadius: 6 },
};

const compactBtnStyles = {
  root: {
    backgroundColor: THEME.background,
    border: `1px solid ${THEME.border}`,
    borderRadius: 6,
    height: 32,
    fontSize: 12,
    fontWeight: '500' as const,
    padding: '0 12px',
    color: THEME.textSecondary,
  },
  rootHovered: { borderColor: THEME.primary, backgroundColor: THEME.primaryLight },
};

const primaryBtnStyles = {
  root: {
    backgroundColor: THEME.primary,
    border: 'none',
    borderRadius: 6,
    height: 32,
    fontSize: 12,
    fontWeight: '500' as const,
    padding: '0 12px',
    color: '#FFFFFF',
  },
  rootHovered: { backgroundColor: THEME.primaryDark },
};

const iconBtnStyles = {
  root: { height: 32, width: 32, color: THEME.textSecondary },
  rootHovered: { color: THEME.primary, backgroundColor: THEME.primaryLight },
};

function getClassNames(compact: boolean) {
  return mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      backgroundColor: THEME.background,
      fontFamily: THEME.fontFamily,
      borderRadius: 4,
      border: `1px solid ${THEME.borderLight}`,
      overflow: 'hidden',
      boxShadow: THEME.shadow,
    },
    // ── Toolbar ──
    toolbar: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      padding: compact ? '6px 16px' : '10px 16px',
      backgroundColor: THEME.background,
      borderBottom: `1px solid ${THEME.border}`,
      flexWrap: 'wrap',
      gap: 8,
      minHeight: compact ? 40 : 48,
    },
    toolbarLeft: { display: 'flex', alignItems: 'center', gap: 8 },
    toolbarRight: { display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' },
    toolbarTitle: {
      fontWeight: 600,
      fontSize: compact ? 14 : 16,
      color: THEME.textPrimary,
    },
    toolbarSubtitle: {
      fontSize: 13,
      color: THEME.textSecondary,
      marginLeft: 8,
    },
    // ── Filter Row ──
    filterRow: {
      display: 'flex',
      gap: 8,
      alignItems: 'center',
      flexWrap: 'wrap',
      padding: '6px 16px',
      backgroundColor: THEME.background,
      borderBottom: `1px solid ${THEME.borderLight}`,
      minHeight: 40,
    },
    searchBox: { width: 200, height: 28 },
    dropdownCompact: { width: 140 },
    // ── Content Area ──
    container: { flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column', position: 'relative' },
    tableContainer: { flex: 1, overflow: 'auto', minHeight: 0 },
    // ── Loading ──
    loadingBox: {
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      alignItems: 'center',
      height: '100%',
      minHeight: 200,
      gap: 16,
      color: THEME.textSecondary,
    },
    // ── Empty ──
    emptyBox: {
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      alignItems: 'center',
      height: '100%',
      minHeight: 240,
      color: THEME.textTertiary,
      backgroundColor: THEME.backgroundAlt,
      margin: 16,
      borderRadius: 12,
      border: `1px dashed ${THEME.border}`,
      padding: 32,
      gap: 12,
    },
    emptyIcon: { fontSize: 48, opacity: 0.4, color: THEME.textTertiary },
    emptyTitle: { fontSize: 16, fontWeight: 600, color: THEME.textPrimary },
    emptySubtitle: { fontSize: 13, color: THEME.textSecondary, maxWidth: 360, textAlign: 'center' },
    // ── Pagination ──
    paginationBar: {
      padding: '8px 16px',
      borderTop: `1px solid ${THEME.borderLight}`,
      backgroundColor: THEME.background,
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
    },
    paginationText: { fontSize: 12, color: THEME.textSecondary },
    paginationCount: { fontWeight: 600, color: THEME.textPrimary },
    // ── Error ──
    errorBox: { padding: 16 },
    // ── Card View (Mobile) ──
    cardContainer: { display: 'flex', flexDirection: 'column', gap: 10, padding: 12, overflowY: 'auto', flex: 1 },
    card: {
      backgroundColor: THEME.background,
      border: `1px solid ${THEME.border}`,
      borderRadius: 10,
      padding: 16,
      cursor: 'pointer',
      boxShadow: THEME.shadow,
      selectors: {
        '&:hover': { boxShadow: '0 4px 12px rgba(0,0,0,0.08)', borderColor: THEME.primary },
      },
    },
    cardHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 10 },
    cardTitle: { fontWeight: 600, fontSize: 15, color: THEME.textPrimary },
    cardBadge: {
      display: 'inline-flex',
      alignItems: 'center',
      padding: '2px 10px',
      borderRadius: 20,
      fontSize: 10,
      fontWeight: 600,
      whiteSpace: 'nowrap',
      marginLeft: 8,
    },
    cardFields: { display: 'flex', flexDirection: 'column', gap: 6 },
    cardField: { display: 'flex', justifyContent: 'space-between', fontSize: 12 },
    cardLabel: { color: THEME.textSecondary },
    cardValue: { color: THEME.textPrimary, fontWeight: 500, textAlign: 'right' as const, maxWidth: '60%', wordBreak: 'break-word' as const },
    // ── Info Bar ──
    infoBar: {
      display: 'flex',
      alignItems: 'center',
      gap: 16,
      padding: '6px 16px',
      backgroundColor: THEME.background,
      borderBottom: `1px solid ${THEME.borderLight}`,
      fontSize: 12,
      color: THEME.textSecondary,
    },
    infoDot: { width: 8, height: 8, borderRadius: '50%', display: 'inline-block' },
    infoLabel: { fontWeight: 600 },
    // ── Column Chooser Panel ──
    columnChooserItem: {
      display: 'flex',
      alignItems: 'center',
      gap: 8,
      padding: '6px 0',
      borderBottom: `1px solid ${THEME.borderLight}`,
    },
    columnChooserLabel: { fontSize: 13, color: THEME.textPrimary },
  });
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

const badgeStyle = (status: string): React.CSSProperties => {
  switch (status?.toLowerCase()) {
    case 'approved': case 'active': case 'paid': case 'issued': case 'complete': case 'ok': case 'available': case 'yes':
      return { backgroundColor: 'rgba(16,124,16,0.08)', color: THEME.success };
    case 'rejected': case 'inactive': case 'blocked': case 'suspended': case 'blacklisted': case 'no':
      return { backgroundColor: 'rgba(209,52,56,0.08)', color: THEME.danger };
    case 'pending': case 'draft': case 'in transit': case 'low': case 'under review': case 'reserved':
      return { backgroundColor: 'rgba(242,200,17,0.12)', color: '#8A6B00' };
    case 'submitted': case 'completed': case 'awarded':
      return { backgroundColor: 'rgba(0,120,212,0.08)', color: THEME.info };
    default:
      return { backgroundColor: '#F0F0F0', color: '#616161' };
  }
};

function formatCellValue(item: IListItem, fieldName: string): string {
  const v = item[fieldName];
  if (v == null) return '';
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    // SharePoint lookup objects come as { Title, ID, Name, etc. }
    return String(
      o.Title || o.Name || o.Material_Name || o.Project_Code ||
      o.Vendor_Name || o.Vendor || o.Code || o.Project_Name ||
      o.Material_Code || o.PO_Number || o.Approver || ''
    );
  }
  return String(v);
}

function formatDateValue(v: unknown): string {
  if (!v) return '';
  try {
    return new Date(String(v)).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch {
    return String(v);
  }
}

function detectStatusField(item: IListItem): string {
  return String(
    item.Status || item.Approval_Status || item.Vendor_Status || item.Project_Status ||
    item.Payment_Status || item.Active || item.Approval_Status || ''
  );
}

function detectTitleField(item: IListItem): string {
  return String(
    item.Title || item.Project_Code || item.Material_Code || item.Material_Name ||
    item.Vendor_Code || item.Vendor_Name || item.Project_Name || item.WarehouseCode || 'Record'
  );
}

// ─── Enterprise DataGrid Component ──────────────────────────────────────────

const EnterpriseDataGrid: React.FC<IDataGridProps> = (props) => {
  const {
    listName, columns, filterQuery, expandQuery, pageSize = 10,
    spHttpClient, pageContext, onRowSelected, onRowDoubleClick,
    title, showBackButton = false, showSearch = true, showRefresh = true,
    showNewRecord = false, showExport = true, onBack, onRefresh, onNewRecord,
    newRecordLabel = '+ New Record', customFilters, statusFilterOptions,
    statusFilterValue, onStatusFilterChange,
    showDateFilter, dateFieldLabel, dateStart, dateEnd, onDateChange,
    sortOptions, sortField, onSortChange,
    showDensityToggle = true,
    showColumnChooser = true,
    showColumnFilters = false,
    className,
  } = props;

  // ── State ──
  const [items, setItems] = React.useState<IListItem[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IListItem[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [currentPage, setCurrentPage] = React.useState(1);
  const [storePageSize] = React.useState(pageSize);
  const [totalItems, setTotalItems] = React.useState(0);
  const [searchQuery, setSearchQuery] = React.useState('');
  const [isMobile, setIsMobile] = React.useState(false);
  const [densityMode, setDensityMode] = React.useState<'comfortable' | 'compact'>('compact');
  const [visibleColumnKeys, setVisibleColumnKeys] = React.useState<Set<string>>(
    new Set(columns.map(c => c.key))
  );
  const [isColumnChooserOpen, setIsColumnChooserOpen] = React.useState(false);
  const [sortColumn, setSortColumn] = React.useState<string | undefined>(undefined);
  const [sortDescending, setSortDescending] = React.useState(false);
  const selectionRef = React.useRef<Selection | null>(null);

  const s = getClassNames(densityMode === 'compact');

  // ── Responsive Check ──
  React.useEffect(() => {
    const check = (): void => setIsMobile(window.innerWidth < 768);
    check();
    window.addEventListener('resize', check);
    return () => window.removeEventListener('resize', check);
  }, []);

  // ── Services ──
  const svc = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  // ── Selection ──
  React.useEffect(() => {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        const sel = selectionRef.current?.getSelection();
        if (sel && sel.length > 0 && onRowSelected) {
          onRowSelected(sel[0] as IListItem);
        }
      },
    });
    return () => { selectionRef.current = null; };
  }, [onRowSelected]);

  // ── Data Fetching ──
  const fetchData = React.useCallback(async (page: number = 1) => {
    try {
      setIsLoading(true);
      setError(undefined);

      const skip = (page - 1) * storePageSize;
      const [fetchedItems, total] = await Promise.all([
        svc.getListData(listName, filterQuery, storePageSize, skip, expandQuery),
        svc.getListItemCount(listName, filterQuery),
      ]);

      setItems(fetchedItems);
      setFilteredItems(fetchedItems);
      setTotalItems(total);
      setCurrentPage(page);
      setSortColumn(undefined);
      setSortDescending(false);
    } catch (err: any) {
      setError(err?.message || 'Failed to load data');
    } finally {
      setIsLoading(false);
    }
  }, [listName, filterQuery, expandQuery, svc, storePageSize]);

  React.useEffect(() => {
    fetchData(1).catch(() => undefined);
  }, [listName, filterQuery, expandQuery]);

  // ── Local Search ──
  React.useEffect(() => {
    if (!searchQuery.trim()) {
      setFilteredItems(items);
      return;
    }
    const q = searchQuery.toLowerCase();
    setFilteredItems(
      items.filter(item =>
        columns.some(col => {
          if (!visibleColumnKeys.has(col.key)) return false;
          const v = item[col.fieldName];
          if (v == null) return false;
          if (typeof v === 'object') {
            return Object.values(v as Record<string, unknown>)
              .some(x => typeof x === 'string' && x.toLowerCase().includes(q));
          }
          return String(v).toLowerCase().includes(q);
        })
      )
    );
  }, [searchQuery, items, columns, visibleColumnKeys]);

  // ── Sorting Handler ──
  const handleSort = React.useCallback((col: IDataGridColumn): void => {
    const isDesc = col.key === sortColumn ? !sortDescending : false;
    setSortColumn(col.key);
    setSortDescending(isDesc);

    const sorted = [...filteredItems].sort((a, b) => {
      const av = a[col.fieldName], bv = b[col.fieldName];
      if (av === bv || (av == null && bv == null)) return 0;
      if (av == null) return isDesc ? -1 : 1;
      if (bv == null) return isDesc ? 1 : -1;
      const cmp = typeof av === 'number' && typeof bv === 'number'
        ? av - bv
        : String(av).localeCompare(String(bv));
      return isDesc ? -cmp : cmp;
    });
    setFilteredItems(sorted);
  }, [filteredItems, sortColumn, sortDescending]);

  // ── Pagination ──
  const handlePage = React.useCallback((page: number): void => {
    fetchData(page).catch(() => undefined);
    setCurrentPage(page);
  }, [fetchData]);

  const totalPages = Math.ceil(totalItems / storePageSize);

  // ── Export ──
  const handleExport = React.useCallback((): void => {
    if (filteredItems.length === 0) return;
    const exportColumns = columns.filter(col => visibleColumnKeys.has(col.key));
    exportToCsv(
      filteredItems as Record<string, unknown>[],
      exportColumns.map(c => ({ key: c.fieldName, label: c.name })),
      `${listName || 'export'}-${new Date().toISOString().slice(0, 10)}.csv`
    );
  }, [filteredItems, columns, visibleColumnKeys, listName]);

  // ── Refresh ──
  const handleRefresh = React.useCallback((): void => {
    onRefresh?.();
    fetchData(1).catch(() => undefined);
  }, [onRefresh, fetchData]);

  // ── Column Chooser ──
  const toggleColumnVisibility = React.useCallback((key: string): void => {
    setVisibleColumnKeys(prev => {
      const next = new Set(prev);
      if (next.has(key)) {
        // Don't allow hiding all columns
        if (next.size > 1) next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  }, []);

  // ── Column Filtering (per-column dropdown filters) ──
  const [columnFilters, setColumnFilters] = React.useState<Record<string, string>>({});

  const handleColumnFilter = React.useCallback((fieldName: string, value: string): void => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (value === '' || value === 'all') {
        delete next[fieldName];
      } else {
        next[fieldName] = value;
      }
      return next;
    });
  }, []);

  // Apply column filters
  React.useEffect(() => {
    let result = [...items];
    const filterKeys = Object.keys(columnFilters);
    if (filterKeys.length > 0) {
      result = result.filter(item =>
        filterKeys.every(key => {
          const filterVal = columnFilters[key];
          if (!filterVal) return true;
          return String(item[key] || '').toLowerCase() === filterVal.toLowerCase();
        })
      );
    }
    // Re-apply search on filtered result
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      result = result.filter(item =>
        columns.some(col => {
          if (!visibleColumnKeys.has(col.key)) return false;
          const v = item[col.fieldName];
          if (v == null) return false;
          return String(v).toLowerCase().includes(q);
        })
      );
    }
    setFilteredItems(result);
  }, [columnFilters, items, searchQuery, columns, visibleColumnKeys]);

  // ── Compute unique values for column filters ──
  const columnFilterOptions = React.useMemo(() => {
    if (!showColumnFilters) return {};
    const options: Record<string, IDropdownOption[]> = {};
    columns.forEach(col => {
      const vals = new Set<string>();
      items.forEach(item => {
        const v = String(item[col.fieldName] || '');
        if (v) vals.add(v);
      });
      options[col.fieldName] = [
        { key: 'all', text: `All ${col.name}` },
        ...Array.from(vals).sort().map(v => ({ key: v, text: v })),
      ];
    });
    return options;
  }, [showColumnFilters, columns, items]);

  // ── Build Display Columns ──
  const displayColumns = React.useMemo(() => {
    return columns
      .filter(col => visibleColumnKeys.has(col.key))
      .map(col => ({
        ...col,
        isSorted: col.key === sortColumn,
        isSortedDescending: col.key === sortColumn ? sortDescending : false,
        onRender: col.onRender || ((item: IListItem) => {
          const v = item[col.fieldName];
          // Null / undefined
          if (v == null) return <span style={{ color: THEME.textTertiary }}>—</span>;

          // SharePoint lookup / person: try the most common object shapes
          if (typeof v === 'object') {
            const o = v as Record<string, unknown>;
            const text = String(
              o.Title || o.Name || o.Material_Name || o.Project_Code ||
              o.Vendor_Name || o.Vendor || o.Code || o.Project_Name ||
              o.Material_Code || o.PO_Number || o.Approver || o.EMail ||
              o.Email || o.Id || o.ID || ''
            );
            // If it's a person field, show with person icon
            if (o.EMail || o.Email || o.Title) {
              return (
                <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{
                    width: 24, height: 24, borderRadius: '50%',
                    backgroundColor: THEME.primaryLight, color: THEME.primary,
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    fontSize: 11, fontWeight: 600, flexShrink: 0,
                  }}>
                    {(text.charAt(0) || '?').toUpperCase()}
                  </span>
                  <span>{text}</span>
                </span>
              );
            }
            return <span>{text || `#${o.ID || item.ID || ''}`}</span>;
          }

          // Boolean: render as tag
          if (typeof v === 'boolean') {
            return (
              <span style={{
                display: 'inline-block', padding: '1px 10px', borderRadius: 12,
                fontSize: 12, fontWeight: 500,
                backgroundColor: v ? 'rgba(16,124,16,0.08)' : 'rgba(209,52,56,0.08)',
                color: v ? THEME.success : THEME.danger,
              }}>
                {v ? 'Yes' : 'No'}
              </span>
            );
          }

          // Number: right-align
          if (typeof v === 'number') {
            return <span style={{ fontVariantNumeric: 'tabular-nums' }}>{v.toLocaleString()}</span>;
          }

          // Date fields: auto-format
          if (col.fieldName === 'Created' || col.fieldName === 'Modified' ||
              col.fieldName.endsWith('_Date') || col.fieldName.endsWith('Date')) {
            return <span style={{ color: THEME.textSecondary }}>{formatDateValue(v)}</span>;
          }

          return <span>{String(v)}</span>;
        }) as (item: any) => React.ReactNode,
      }));
  }, [columns, visibleColumnKeys, sortColumn, sortDescending]);

  // ── Card Renderer ──
  const renderCard = React.useCallback((item: IListItem): React.ReactNode => {
    const status = detectStatusField(item);
    const ttl = detectTitleField(item);
    const fields = displayColumns
      .filter(c => !['ID', 'Title', 'Status', 'Approval_Status', 'Vendor_Status', 'Project_Status', 'Payment_Status', 'Active'].includes(c.fieldName))
      .slice(0, 5);

    return (
      <div
        key={item.ID}
        className={s.card}
        onClick={() => onRowDoubleClick?.(item)}
      >
        <div className={s.cardHeader}>
          <Text className={s.cardTitle} styles={{ root: { flex: 1 } }}>{ttl}</Text>
          {status && (
            <span className={s.cardBadge} style={badgeStyle(status)}>{status}</span>
          )}
        </div>
        <div className={s.cardFields}>
          {fields.map(f => {
            const v = item[f.fieldName];
            if (v == null || v === '') return null;
            return (
              <div key={f.fieldName} className={s.cardField}>
                <Text className={s.cardLabel}>{f.name}</Text>
                <Text className={s.cardValue}>
                  {typeof v === 'object'
                    ? formatCellValue(item, f.fieldName)
                    : String(v)}
                </Text>
              </div>
            );
          })}
        </div>
      </div>
    );
  }, [displayColumns, onRowDoubleClick, s]);

  // ── DetailsList Styles ──
  const detailsListStyles = React.useMemo((): any => ({
    root: {
      selectors: {
        // ── HEADER: White bg, no uppercase, 13px Segoe UI Semi Bold ──
        '.ms-DetailsHeader': {
          borderBottom: `1px solid ${THEME.borderLight}`,
          paddingTop: 0,
          paddingBottom: 0,
          backgroundColor: THEME.headerBg,
          height: 40,
          lineHeight: '40px',
          selectors: {
            ':hover': { backgroundColor: THEME.headerBg },
          },
        },
        '.ms-DetailsHeader-cell': {
          height: 40,
          lineHeight: '40px',
        },
        '.ms-DetailsHeader-cellTitle': {
          fontWeight: 600,
          fontSize: 13,
          textTransform: 'none',
          letterSpacing: 'normal',
          color: THEME.textPrimary,
          fontFamily: THEME.fontFamily,
          height: 40,
          lineHeight: '40px',
          paddingLeft: 12,
          paddingRight: 8,
        },
        '.ms-DetailsHeader-cell:hover .ms-DetailsHeader-cellTitle': {
          color: THEME.primary,
        },
        '.ms-DetailsHeader-cell.is-sorted .ms-DetailsHeader-cellTitle': {
          color: THEME.primary,
        },
        // ── SORTING INDICATOR ──
        '.ms-DetailsHeader-sortIcon': {
          color: THEME.primary,
        },
        '.ms-DetailsHeader-cell.is-sorted .ms-DetailsHeader-sortIcon': {
          color: THEME.primary,
        },
        // ── ROWS: Clean alternating white, no uppercase ──
        '.ms-DetailsRow': {
          borderBottom: `1px solid ${THEME.borderLight}`,
          fontSize: 13,
          minHeight: 38,
          lineHeight: 'normal',
          selectors: {
            '&:hover': { backgroundColor: THEME.primaryLight, cursor: 'pointer' },
            '&:nth-child(even)': { backgroundColor: THEME.backgroundAlt },
            '&:nth-child(even):hover': { backgroundColor: THEME.primaryLight },
            '&.is-selected': {
              backgroundColor: 'rgba(0,120,212,0.08)',
              borderBottom: `1px solid ${THEME.primary}`,
            },
          },
        },
        '.ms-DetailsRow-fields': {
          fontSize: 13,
          fontFamily: THEME.fontFamily,
        },
        '.ms-DetailsRow-cell': {
          fontSize: 13,
          fontFamily: THEME.fontFamily,
          color: THEME.textPrimary,
          paddingLeft: 12,
          paddingRight: 8,
          lineHeight: 'normal',
          display: 'flex',
          alignItems: 'center',
        },
        '.ms-FocusZone': { minHeight: 0 },
        // ── CHECKBOX HIDDEN ──
        '.ms-DetailsHeader-cell.is-check': { display: 'none' },
        '.ms-DetailsRow-cell.is-check': { display: 'none' },
      },
    },
  }), [densityMode]);

  // ── Render ──
  const hasFilters = showSearch || statusFilterOptions || customFilters || showDateFilter || sortOptions || showDensityToggle || showColumnChooser;

  return (
    <div className={`${s.root} ${className || ''}`.trim()}>
      {/* Error */}
      {error && (
        <div className={s.errorBox}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>
        </div>
      )}

      {/* Toolbar */}
      {(title || showBackButton || showRefresh || showNewRecord || showExport || showDensityToggle || showColumnChooser) && (
        <div className={s.toolbar}>
          <div className={s.toolbarLeft}>
            {showBackButton && (
              <IconButton iconProps={{ iconName: 'Back' }} title="Back" onClick={onBack} styles={iconBtnStyles} />
            )}
            {title && (
              <>
                <Text className={s.toolbarTitle}>{title}</Text>
                {totalItems > 0 && (
                  <Text className={s.toolbarSubtitle}>
                    {totalItems} record{totalItems !== 1 ? 's' : ''}
                  </Text>
                )}
              </>
            )}
          </div>
          <div className={s.toolbarRight}>
            {showExport && filteredItems.length > 0 && (
              <DefaultButton
                text="Export"
                iconProps={{ iconName: 'ExcelDocument' }}
                onClick={handleExport}
                styles={compactBtnStyles}
              />
            )}
            {showNewRecord && (
              <PrimaryButton
                text={newRecordLabel}
                iconProps={{ iconName: 'Add' }}
                onClick={onNewRecord}
                styles={primaryBtnStyles}
              />
            )}
            {showDensityToggle && (
              <TooltipHost
                content={densityMode === 'compact' ? 'Switch to comfortable view' : 'Switch to compact view'}
                directionalHint={DirectionalHint.bottomCenter}
              >
                <IconButton
                  iconProps={{ iconName: densityMode === 'compact' ? 'GridViewMedium' : 'BulletedList' }}
                  title="Toggle density"
                  onClick={() => setDensityMode(prev => prev === 'compact' ? 'comfortable' : 'compact')}
                  styles={iconBtnStyles}
                />
              </TooltipHost>
            )}
            {showColumnChooser && (
              <TooltipHost content="Choose columns" directionalHint={DirectionalHint.bottomCenter}>
                <IconButton
                  iconProps={{ iconName: 'BulletedList' }}
                  title="Column chooser"
                  onClick={() => setIsColumnChooserOpen(true)}
                  styles={iconBtnStyles}
                />
              </TooltipHost>
            )}
            {showRefresh && (
              <TooltipHost content="Refresh data" directionalHint={DirectionalHint.bottomCenter}>
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  title="Refresh"
                  onClick={handleRefresh}
                  styles={iconBtnStyles}
                />
              </TooltipHost>
            )}
          </div>
        </div>
      )}

      {/* Info Bar */}
      {totalItems > 0 && !isLoading && (
        <div className={s.infoBar}>
          <span>
            Showing <strong className={s.paginationCount}>{filteredItems.length}</strong>
            {filteredItems.length < totalItems && ` of ${totalItems}`} records
          </span>
          {sortColumn && (
            <span>
              <span className={s.infoDot} style={{ backgroundColor: THEME.info }} />{' '}
              Sorted by {columns.find(c => c.key === sortColumn)?.name || sortColumn}
              {sortDescending ? ' ↓' : ' ↑'}
            </span>
          )}
          {searchQuery && (
            <span>
              <span className={s.infoDot} style={{ backgroundColor: THEME.warning }} />{' '}
              Filtered: "{searchQuery}"
            </span>
          )}
        </div>
      )}

      {/* Filter Row */}
      {hasFilters && (
        <div className={s.filterRow}>
          {showSearch && (
            <SearchBox
              placeholder="Search all columns..."
              value={searchQuery}
              onChange={(_, v) => setSearchQuery(v || '')}
              onClear={() => setSearchQuery('')}
              styles={{
                root: { width: 200, height: 28 },
                field: { fontSize: 12 },
              }}
            />
          )}
          {statusFilterOptions && onStatusFilterChange && (
            <Dropdown
              placeholder="All Status"
              options={statusFilterOptions}
              selectedKey={statusFilterValue}
              onChange={(_, opt) => onStatusFilterChange((opt?.key as string) || '')}
              styles={{
                root: { width: 140 },
                dropdown: { borderRadius: 6, border: `1px solid ${THEME.border}`, height: 28, lineHeight: 26, fontSize: 12 },
              }}
            />
          )}
          {sortOptions && onSortChange && (
            <Dropdown
              placeholder="Sort by"
              options={sortOptions}
              selectedKey={sortField}
              onChange={(_, opt) => onSortChange((opt?.key as string) || '')}
              styles={{
                root: { width: 140 },
                dropdown: { borderRadius: 6, border: `1px solid ${THEME.border}`, height: 28, lineHeight: 26, fontSize: 12 },
              }}
            />
          )}
          {showDateFilter && (
            <>
              <DatePicker
                placeholder="From"
                value={dateStart}
                onSelectDate={(d) => onDateChange?.(d || undefined, dateEnd)}
                styles={dateStyles}
                textField={{ styles: { root: { height: 28 }, field: { fontSize: 12, padding: '0 8px' } } }}
              />
              <DatePicker
                placeholder="To"
                value={dateEnd}
                onSelectDate={(d) => onDateChange?.(dateStart, d || undefined)}
                styles={dateStyles}
                textField={{ styles: { root: { height: 28 }, field: { fontSize: 12, padding: '0 8px' } } }}
              />
            </>
          )}
          {customFilters}
        </div>
      )}

      {/* Per-Column Quick Filters */}
      {showColumnFilters && displayColumns.length > 0 && (
        <div className={s.filterRow} style={{ backgroundColor: THEME.background, padding: '4px 16px', borderBottom: `1px solid ${THEME.borderLight}` }}>
          {displayColumns.slice(0, 6).map(col => {
            const options = columnFilterOptions[col.fieldName];
            if (!options || options.length <= 1) return null;
            return (
              <Dropdown
                key={col.fieldName}
                placeholder={col.name}
                options={options}
                selectedKey={columnFilters[col.fieldName] || 'all'}
                onChange={(_, opt) => handleColumnFilter(col.fieldName, (opt?.key as string) || 'all')}
                styles={{
                  root: { width: 130 },
                  dropdown: { borderRadius: 6, border: `1px solid ${THEME.border}`, height: 24, lineHeight: 22, fontSize: 11 },
                }}
              />
            );
          })}
        </div>
      )}

      {/* Content Area */}
      <div className={s.container} role="region" aria-label="Data grid content">
        {isLoading ? (
          <div className={s.loadingBox}>
            <Spinner size={SpinnerSize.medium} />
            <Text variant="medium" style={{ color: THEME.textSecondary }}>Loading data...</Text>
          </div>
        ) : filteredItems.length === 0 ? (
          <div className={s.emptyBox}>
            <Icon iconName="ClipboardList" className={s.emptyIcon} />
            <Text className={s.emptyTitle}>No records found</Text>
            <Text className={s.emptySubtitle}>
              {searchQuery || Object.keys(columnFilters).length > 0
                ? 'Try adjusting your search or filter criteria to find what you\'re looking for.'
                : 'No data is available yet. Add a new record to get started.'}
            </Text>
          </div>
        ) : (
          <>
            {isMobile ? (
              <div className={s.cardContainer}>
                {filteredItems.map(item => renderCard(item))}
              </div>
            ) : (
              <div className={s.tableContainer}>
                <DetailsList
                  items={filteredItems}
                  columns={displayColumns}
                  selectionMode={SelectionMode.single}
                  selection={selectionRef.current!}
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  constrainMode={ConstrainMode.unconstrained}
                  isHeaderVisible={true}
                  checkboxVisibility={CheckboxVisibility.hidden}
                  onColumnHeaderClick={(_, col) => col && handleSort(col as IDataGridColumn)}
                  onItemInvoked={(item) => onRowDoubleClick?.(item)}
                  styles={detailsListStyles}
                  compact={densityMode === 'compact'}
                />
              </div>
            )}
            {/* Pagination */}
            {totalPages > 1 && (
              <div className={s.paginationBar}>
                <Text className={s.paginationText}>
                  Page <strong className={s.paginationCount}>{currentPage}</strong> of {totalPages} &middot;{' '}
                  {totalItems.toLocaleString()} total records
                </Text>
                <Pagination
                  currentPage={currentPage}
                  totalPages={totalPages}
                  onChange={handlePage}
                  limiter={3}
                />
              </div>
            )}
          </>
        )}
      </div>

      {/* Column Chooser Panel */}
      <Panel
        isOpen={isColumnChooserOpen}
        onDismiss={() => setIsColumnChooserOpen(false)}
        type={PanelType.smallFixedFar}
        headerText="Choose Columns"
        closeButtonAriaLabel="Close"
        styles={{
          main: { borderTopLeftRadius: 12, borderBottomLeftRadius: 12, boxShadow: '-4px 0 24px rgba(0,0,0,0.08)' },
          headerText: { fontWeight: 600, fontSize: 16 },
        }}
      >
        <div style={{ padding: '8px 0' }}>
          <Text variant="small" style={{ color: THEME.textTertiary, display: 'block', marginBottom: 12 }}>
            Toggle columns to show or hide them in the table.
          </Text>
          {columns.map(col => (
            <div key={col.key} className={s.columnChooserItem}>
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', flex: 1 }}>
                <input
                  type="checkbox"
                  checked={visibleColumnKeys.has(col.key)}
                  onChange={() => toggleColumnVisibility(col.key)}
                  style={{ accentColor: THEME.primary }}
                />
                <span className={s.columnChooserLabel}>{col.name}</span>
              </label>
              <Text variant="small" style={{ color: THEME.textTertiary, fontSize: 11 }}>
                {col.minWidth}px
              </Text>
            </div>
          ))}
        </div>
      </Panel>
    </div>
  );
};

export default EnterpriseDataGrid;
