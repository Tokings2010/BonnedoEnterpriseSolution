import * as React from 'react';
import {
  DetailsList, DetailsListLayoutMode, Selection, IColumn, SelectionMode,
  Spinner, SpinnerSize, MessageBar, MessageBarType, mergeStyleSets,
  Text, Icon, SearchBox, IconButton, PrimaryButton, DefaultButton,
  Dropdown, IDropdownOption, DatePicker, IDatePickerStyles,
} from '@fluentui/react';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';
import { exportToCsv } from './TagRenderer';

export interface IDataGridColumn extends IColumn {
  fieldName: string; name: string; minWidth: number; maxWidth?: number;
  isResizable?: boolean; isSorted?: boolean; isSortedDescending?: boolean;
}
export interface IDataGridProps {
  listName: string; columns: IDataGridColumn[]; filterQuery?: string;
  expandQuery?: string; pageSize?: number;
  spHttpClient: SPHttpClient; pageContext: PageContext;
  onRowSelected?: (record: IListItem) => void;
  onRowDoubleClick?: (record: IListItem) => void;
  title?: string; showBackButton?: boolean; showSearch?: boolean;
  showRefresh?: boolean; showNewRecord?: boolean; showExport?: boolean;
  onBack?: () => void; onRefresh?: () => void; onNewRecord?: () => void;
  newRecordLabel?: string; customFilters?: React.ReactNode;
  statusFilterOptions?: IDropdownOption[]; statusFilterValue?: string;
  onStatusFilterChange?: (value: string) => void;
  /** Enable date range filter */
  showDateFilter?: boolean;
  /** Date filter label */
  dateFieldLabel?: string;
  /** Current start date */
  dateStart?: Date;
  /** Current end date */
  dateEnd?: Date;
  /** Date range change handler */
  onDateChange?: (start: Date | undefined, end: Date | undefined) => void;
  className?: string;
  /** Show sort dropdown */
  sortOptions?: IDropdownOption[];
  sortField?: string;
  onSortChange?: (field: string) => void;
}

interface IState {
  items: IListItem[]; filteredItems: IListItem[]; columns: IDataGridColumn[];
  isLoading: boolean; error: string | undefined; currentPage: number;
  pageSize: number; totalItems: number; searchQuery: string; isMobile: boolean;
}

const dateStyles: Partial<IDatePickerStyles> = {
  root: { maxWidth: 140, minWidth: 130 },
  textField: { border: '1px solid #E8ECF0', borderRadius: 6 },
};

const compactBtnStyles = {
  root: { backgroundColor: '#FFFFFF', border: '1px solid #E8ECF0', borderRadius: 6, height: 32, fontSize: 12, fontWeight: '500' as const, padding: '0 10px' },
  rootHovered: { borderColor: '#00B894', backgroundColor: 'rgba(0,184,148,0.04)' },
};

const EnhancedDataGrid: React.FC<IDataGridProps> = (props) => {
  const {
    listName, columns, filterQuery, expandQuery, pageSize = 10,
    spHttpClient, pageContext, onRowSelected, onRowDoubleClick,
    title, showBackButton = false, showSearch = true, showRefresh = true,
    showNewRecord = false, showExport = true, onBack, onRefresh, onNewRecord,
    newRecordLabel = '+ New Record', customFilters, statusFilterOptions,
    statusFilterValue, onStatusFilterChange,
    showDateFilter, dateFieldLabel, dateStart, dateEnd, onDateChange,
    sortOptions, sortField, onSortChange,
    className,
  } = props;

  const [st, setSt] = React.useState<IState>({
    items: [], filteredItems: [], columns, isLoading: true, error: undefined,
    currentPage: 1, pageSize, totalItems: 0, searchQuery: '', isMobile: false,
  });

  React.useEffect(() => {
    const chk = (): void => setSt(p => ({ ...p, isMobile: window.innerWidth < 768 }));
    chk(); window.addEventListener('resize', chk);
    return () => window.removeEventListener('resize', chk);
  }, []);

  React.useEffect(() => {
    if (!st.searchQuery.trim()) { setSt(p => ({ ...p, filteredItems: st.items })); return; }
    const q = st.searchQuery.toLowerCase();
    setSt(p => ({
      ...p,
      filteredItems: st.items.filter(item =>
        st.columns.some(col => {
          const v = item[col.fieldName];
          if (v == null) return false;
          if (typeof v === 'object') return Object.values(v as Record<string, unknown>)
            .some(x => typeof x === 'string' && x.toLowerCase().includes(q));
          return String(v).toLowerCase().includes(q);
        })
      ),
    }));
  }, [st.searchQuery, st.items, st.columns]);

  const svc = React.useMemo(() => new SharePointService(spHttpClient, pageContext), [spHttpClient, pageContext]);
  const sel = React.useRef(new Selection({
    onSelectionChanged: () => {
      const s = sel.current.getSelection();
      if (s.length > 0 && onRowSelected) onRowSelected(s[0] as IListItem);
    }
  }));

  const fetchData = React.useCallback(async (page = 1) => {
    try {
      setSt(p => ({ ...p, isLoading: true, error: undefined }));
      const skip = (page - 1) * st.pageSize;
      const items = await svc.getListData(listName, filterQuery, st.pageSize, skip, expandQuery);
      const total = await svc.getListItemCount(listName, filterQuery);
      setSt(p => ({ ...p, items, filteredItems: items, totalItems: total, currentPage: page, isLoading: false }));
    } catch (err) {
      setSt(p => ({ ...p, isLoading: false, error: err instanceof Error ? err.message : 'Failed' }));
    }
  }, [listName, filterQuery, expandQuery, svc, st.pageSize]);

  React.useEffect(() => { fetchData(1).catch(() => undefined); }, [listName, filterQuery, expandQuery]);

  const handleSort = (col: IDataGridColumn): void => {
    const desc = !col.isSortedDescending;
    const cols = st.columns.map(c => ({
      ...c, isSorted: c.key === col.key,
      isSortedDescending: c.key === col.key ? desc : false
    }));
    const sorted = [...st.filteredItems].sort((a, b) => {
      const av = a[col.fieldName], bv = b[col.fieldName];
      if (av === bv) return 0;
      return desc ? (av < bv ? 1 : -1) : av < bv ? -1 : 1;
    });
    setSt(p => ({ ...p, columns: cols, filteredItems: sorted }));
  };

  const handlePage = (n: number): void => {
    fetchData(n).catch(() => undefined);
    sel.current.setAllSelected(false);
  };

  const s = mergeStyleSets({
    root: { display: 'flex', flexDirection: 'column', height: '100%', backgroundColor: '#FFFFFF' },
    toolbar: {
      display: 'flex', justifyContent: 'space-between', alignItems: 'center',
      padding: '8px 16px', backgroundColor: '#FFFFFF', borderBottom: '1px solid #E8ECF0',
      flexWrap: 'wrap', gap: 8, minHeight: 40,
    },
    toolbarLeft: { display: 'flex', alignItems: 'center', gap: 8 },
    toolbarRight: { display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' },
    toolbarTitle: { fontWeight: 600, fontSize: 14, color: '#2C3E50' },
    compactBtn: { border: '1px solid #E8ECF0', borderRadius: 6, height: 28, fontSize: 12, padding: '0 10px', backgroundColor: '#FFFFFF' },
    container: { flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' },
    tableContainer: { flex: 1, overflow: 'auto', minHeight: 0 },
    paginationBar: { padding: '8px 16px', borderTop: '1px solid #E8ECF0', backgroundColor: '#FFFFFF', display: 'flex', justifyContent: 'space-between', alignItems: 'center' },
    paginationText: { fontSize: 11, color: '#7F8C9B' },
    errorBox: { padding: 16 },
    loadingBox: { display: 'flex', justifyContent: 'center', alignItems: 'center', height: 200 },
    emptyBox: { display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', height: 200, color: '#7F8C9B', backgroundColor: '#FAF9FC', border: '1px dashed #E8ECF0', borderRadius: 8, margin: 16, padding: 32 },
    cardContainer: { display: 'flex', flexDirection: 'column', gap: 12, padding: 12, overflowY: 'auto' },
    card: { backgroundColor: '#FFFFFF', border: '1px solid #E8ECF0', borderRadius: 8, padding: 16, cursor: 'pointer', boxShadow: '0 1px 3px rgba(0,0,0,0.04)', selectors: { '&:hover': { boxShadow: '0 2px 8px rgba(0,0,0,0.06)' } } },
    cardHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 },
    cardTitle: { fontWeight: 600, fontSize: 15, color: '#2C3E50' },
    cardBadge: { display: 'inline-flex', alignItems: 'center', padding: '2px 10px', borderRadius: 20, fontSize: 10, fontWeight: 600 },
    cardFields: { display: 'flex', flexDirection: 'column', gap: 4 },
    cardField: { display: 'flex', justifyContent: 'space-between', fontSize: 13 },
    cardLabel: { color: '#7F8C9B' },
    cardValue: { color: '#2C3E50', fontWeight: 500 },
    filterRow: {
      display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap',
      padding: '8px 16px', backgroundColor: '#FAF9FC', borderBottom: '1px solid #E8ECF0',
      minHeight: 36,
    },
  });

  const totalPages = Math.ceil(st.totalItems / st.pageSize);
  const hasFilters = showSearch || statusFilterOptions || customFilters || showDateFilter || sortOptions;

  const displayCols = st.columns.map(col => ({
    ...col,
    onRender: col.onRender || ((item: IListItem) => {
      const v = item[col.fieldName];
      if (v && typeof v === 'object') {
        const o = v as Record<string, unknown>;
        return o.Title || o.Name || o.Material_Name || o.Project_Code || o.Vendor_Name || String(v);
      }
      if (v && (col.fieldName === 'Created' || col.fieldName.endsWith('_Date') || col.fieldName.endsWith('Date'))) {
        try { return new Date(String(v)).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); }
        catch { return String(v); }
      }
      return v ?? '';
    }),
  }));

  const badgeStyle = (status: string): React.CSSProperties => {
    switch (status?.toLowerCase()) {
      case 'approved': case 'active': case 'paid': case 'issued': case 'complete': case 'ok':
        return { backgroundColor: 'rgba(0,184,148,0.08)', color: '#00B894' };
      case 'rejected': case 'inactive': case 'blocked':
        return { backgroundColor: 'rgba(231,76,60,0.08)', color: '#E74C3C' };
      case 'pending': case 'draft': case 'in transit': case 'low':
        return { backgroundColor: 'rgba(243,156,18,0.08)', color: '#F39C12' };
      case 'submitted': case 'completed':
        return { backgroundColor: 'rgba(74,144,217,0.08)', color: '#4A90D9' };
      default:
        return { backgroundColor: '#F3F4F6', color: '#6B7280' };
    }
  };

  const handleExport = (): void => {
    exportToCsv(
      st.filteredItems as Record<string, unknown>[],
      st.columns.map(c => ({ key: c.fieldName, label: c.name })),
      listName + '.csv'
    );
  };

  return (
    <div className={(s.root + ' ' + (className || '')).trim()}>
      {st.error && <div className={s.errorBox}><MessageBar messageBarType={MessageBarType.error} isMultiline>{st.error}</MessageBar></div>}
      {(title || showBackButton || showRefresh || showNewRecord || showExport) && (
        <div className={s.toolbar}>
          <div className={s.toolbarLeft}>
            {showBackButton && <IconButton iconProps={{ iconName: 'Back' }} title="Back" onClick={onBack} styles={{ root: { height: 28, width: 28 } }} />}
            {title && <Text className={s.toolbarTitle}>{title}</Text>}
          </div>
          <div className={s.toolbarRight}>
            {showExport && st.filteredItems.length > 0 && (
              <DefaultButton text="Export" iconProps={{ iconName: 'ExcelDocument' }} onClick={handleExport} styles={compactBtnStyles} />
            )}
            {showNewRecord && (
              <PrimaryButton text={newRecordLabel} iconProps={{ iconName: 'Add' }} onClick={onNewRecord}
                styles={{ root: { backgroundColor: '#00B894', border: 'none', borderRadius: 6, height: 28, fontSize: 12, fontWeight: '500' as const, padding: '0 10px' }, rootHovered: { backgroundColor: '#009E80' } }} />
            )}
            {showRefresh && (
              <IconButton iconProps={{ iconName: 'Refresh' }} title="Refresh"
                onClick={() => { onRefresh?.(); if (!onRefresh) fetchData(1).catch(() => undefined); }} styles={{ root: { height: 28, width: 28 } }} />
            )}
          </div>
        </div>
      )}
      {hasFilters && (
        <div className={s.filterRow}>
          {showSearch && (
            <SearchBox placeholder="Search..." value={st.searchQuery}
              onChange={(_, v) => setSt(p => ({ ...p, searchQuery: v || '' }))}
              onClear={() => setSt(p => ({ ...p, searchQuery: '' }))}
              styles={{ root: { width: st.isMobile ? '100%' : 200, height: 28, borderRadius: 6, border: '1px solid #E8ECF0' } }} />
          )}
          {statusFilterOptions && onStatusFilterChange && (
            <Dropdown placeholder="All Status" options={statusFilterOptions} selectedKey={statusFilterValue}
              onChange={(_, opt) => onStatusFilterChange((opt?.key as string) || '')}
              styles={{ root: { width: st.isMobile ? '100%' : 150 }, dropdown: { borderRadius: 6, border: '1px solid #E8ECF0', height: 28, lineHeight: 26, fontSize: 12 } }} />
          )}
          {sortOptions && onSortChange && (
            <Dropdown placeholder="Sort by" options={sortOptions} selectedKey={sortField}
              onChange={(_, opt) => onSortChange((opt?.key as string) || '')}
              styles={{ root: { width: st.isMobile ? '100%' : 140 }, dropdown: { borderRadius: 6, border: '1px solid #E8ECF0', height: 28, lineHeight: 26, fontSize: 12 } }} />
          )}
          {showDateFilter && (
            <>
              <DatePicker placeholder="From" value={dateStart}
                onSelectDate={(d) => onDateChange?.(d || undefined, dateEnd)}
                styles={dateStyles}
                textField={{ styles: { root: { height: 28 }, field: { fontSize: 12, padding: '0 8px' } } }} />
              <DatePicker placeholder="To" value={dateEnd}
                onSelectDate={(d) => onDateChange?.(dateStart, d || undefined)}
                styles={dateStyles}
                textField={{ styles: { root: { height: 28 }, field: { fontSize: 12, padding: '0 8px' } } }} />
            </>
          )}
          {customFilters}
        </div>
      )}
      <div className={s.container}>
        {st.isLoading ? (
          <div className={s.loadingBox}><Spinner size={SpinnerSize.medium} label="Loading..." /></div>
        ) : st.filteredItems.length === 0 ? (
          <div className={s.emptyBox}>
            <Icon iconName="ClipboardEmpty" style={{ fontSize: 40, marginBottom: 12, opacity: 0.5 }} />
            <Text variant="medium" style={{ color: '#7F8C9B' }}>No items found</Text>
          </div>
        ) : (
          <>
            {st.isMobile ? (
              <div className={s.cardContainer}>
                {st.filteredItems.map(item => {
                  const status = String(item.Status || item.Approval_Status || item.Vendor_Status || item.Project_Status || '');
                  const ttl = String(item.Title || item.Project_Code || item.Material_Code || item.Project_Name || 'Record');
                  const fields = st.columns.filter(c => !['ID','Title','Status','Approval_Status','Vendor_Status','Project_Status'].includes(c.fieldName)).slice(0, 4);
                  return (
                    <div key={item.ID} className={s.card} onClick={() => onRowDoubleClick?.(item)}>
                      <div className={s.cardHeader}>
                        <Text className={s.cardTitle}>{ttl}</Text>
                        {status && <span className={s.cardBadge} style={badgeStyle(status)}>{status}</span>}
                      </div>
                      <div className={s.cardFields}>
                        {fields.map(f => {
                          const v = item[f.fieldName];
                          if (!v) return null;
                          return (
                            <div key={f.fieldName} className={s.cardField}>
                              <Text className={s.cardLabel}>{f.name}</Text>
                              <Text className={s.cardValue}>{typeof v === 'object' ? JSON.stringify(v) : String(v)}</Text>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className={s.tableContainer}>
                <DetailsList items={st.filteredItems} columns={displayCols}
                  selectionMode={SelectionMode.single} selection={sel.current}
                  layoutMode={DetailsListLayoutMode.fixedColumns} isHeaderVisible={true}
                  onColumnHeaderClick={(_, col) => col && handleSort(col as IDataGridColumn)}
                  onItemInvoked={(item) => onRowDoubleClick?.(item)}
                  styles={{
                    root: {
                      selectors: {
                        '.ms-DetailsHeader': {
                          borderBottom: '1px solid #E8ECF0', paddingTop: 4, paddingBottom: 4,
                          backgroundColor: '#F5F6FA', height: 32, lineHeight: 32,
                        },
                        '.ms-DetailsHeader-cellTitle': {
                          fontWeight: 600, fontSize: 11, textTransform: 'uppercase',
                          letterSpacing: '0.02em', color: '#7F8C9B',
                          fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', sans-serif",
                        },
                        '.ms-DetailsHeader-cell:hover .ms-DetailsHeader-cellTitle': { color: '#2C3E50' },
                        '.ms-DetailsRow': {
                          borderBottom: '1px solid #F0F2F5', fontSize: 12,
                          selectors: {
                            '&:hover': { backgroundColor: 'rgba(0,184,148,0.02)', cursor: 'pointer' },
                            '&.is-selected': { backgroundColor: 'rgba(0,184,148,0.04)' },
                          },
                        },
                        '.ms-DetailsRow-fields': { fontSize: 12 },
                        '.ms-DetailsRow-cell': { fontSize: 12, fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', sans-serif" },
                      },
                    },
                  }} />
              </div>
            )}
            {totalPages > 1 && (
              <div className={s.paginationBar}>
                <Text className={s.paginationText}>Page {st.currentPage} of {totalPages} | {st.totalItems} items</Text>
                <Pagination currentPage={st.currentPage} totalPages={totalPages} onChange={handlePage} />
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};

export default EnhancedDataGrid;
