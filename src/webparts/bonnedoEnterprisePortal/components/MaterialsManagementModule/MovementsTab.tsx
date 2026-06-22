import * as React from 'react';
import {
  mergeStyleSets,
  Spinner, SpinnerSize,
  MessageBar, MessageBarType,
  Stack, Text, Icon,
  Dropdown, IDropdownOption,
  DatePicker,
  SearchBox,
  CommandBar, ICommandBarItemProps,
  Link,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../../services/SharePointService';
import { SHAREPOINT_LISTS } from '../../constants/SharePointListNames';
import { IMaterialMasterRecord } from '../../models/DataModels';
import { IQRScanRequestContext } from './QRScanMovePanel';
import EnhancedDataGrid from '../EnhancedDataGrid';
import { IDataGridColumn } from '../EnhancedDataGrid';
import { Tag, exportToCsv } from '../TagRenderer';

export interface IMovementsTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
  onQrScanRequested: (context?: IQRScanRequestContext) => void;
}

interface IMovementItem {
  ID: number;
  Title: string;
  MovementType: string;
  MaterialCode: string;
  MaterialName: string;
  Category: string;
  QRCodeURL: string;
  Qty: number;
  FromLocation: string;
  ToLocation: string;
  ProjectCode: string;
  Note: string;
  Created: string;
}

const MOVEMENT_COLORS: Record<string, { bg: string; color: string }> = {
  GRN: { bg: '#D1FAE5', color: '#065F46' },
  'Transfer Out': { bg: '#DBEAFE', color: '#1E40AF' },
  Issue: { bg: '#FEF3C7', color: '#92400E' },
  Return: { bg: '#E0E7FF', color: '#3730A3' },
  Scrap: { bg: '#FEE2E2', color: '#991B1B' },
};

const TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'all', text: 'All Movement Types' },
  { key: 'GRN', text: 'GRN (Receive)' },
  { key: 'Transfer Out', text: 'Transfer' },
  { key: 'Issue', text: 'Issue' },
  { key: 'Return', text: 'Return' },
  { key: 'Scrap', text: 'Scrap' },
];

interface IMovementsStyles {
  container: string;
  filterRow: string;
  movementPill: string;
  summaryRow: string;
  summaryCard: string;
  cardValue: string;
  cardLabel: string;
  emptyState: string;
}

const getStyles = (): IMovementsStyles =>
  mergeStyleSets({
    container: { display: 'flex', flexDirection: 'column', gap: '16px' },
    filterRow: { display: 'flex', gap: '12px', alignItems: 'flex-end', flexWrap: 'wrap' },
    movementPill: {
      display: 'inline-block',
      padding: '3px 12px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
    },
    summaryRow: { display: 'flex', gap: '16px', flexWrap: 'wrap' },
    summaryCard: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '12px 20px',
      display: 'flex',
      alignItems: 'center',
      gap: '10px',
      minWidth: '120px',
    },
    cardValue: { fontSize: '20px', fontWeight: '700', color: '#1E2532' },
    cardLabel: { fontSize: '12px', color: '#5A6A85' },
    emptyState: { textAlign: 'center', padding: '60px 20px', color: '#5A6A85' },
  });

function escapeCsvValue(value: string | number): string {
  const str = String(value ?? '');
  return `"${str.replace(/"/g, '""')}"`;
}

function exportMovementsToCsv(items: IMovementItem[]): void {
  const headers = ['Date', 'Type', 'Material Code', 'Qty', 'From', 'To', 'Project', 'Note'];
  const rows = items.map((item) => [
    item.Created ? new Date(item.Created).toISOString() : '',
    item.MovementType,
    item.MaterialCode,
    item.Qty,
    item.FromLocation,
    item.ToLocation,
    item.ProjectCode,
    item.Note,
  ]);
  const csv = [headers.join(','), ...rows.map((row) => row.map(escapeCsvValue).join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `inventory-movements-${new Date().toISOString().slice(0, 10)}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
}

const MovementsTab: React.FC<IMovementsTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
  onQrScanRequested,
}) => {
  const [items, setItems] = React.useState<IMovementItem[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IMovementItem[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [selectedType, setSelectedType] = React.useState<string>('all');
  const [searchText, setSearchText] = React.useState('');
  const [startDate, setStartDate] = React.useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = React.useState<Date | undefined>(undefined);

  const classNames = getStyles();

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const fetchMovements = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(null);
    try {
      const [movementRecords, materialRecords] = await Promise.all([
        sharePointService.getInventoryMovements(),
        sharePointService.getMaterialMasterRecords().catch((): IMaterialMasterRecord[] => []),
      ]);
      const materialMap = new Map<string, IMaterialMasterRecord>();
      materialRecords.forEach((material) => {
        const code = material.Material_Code || material.Title || '';
        if (code) {
          materialMap.set(code, material);
          materialMap.set(material.Title, material);
        }
      });
      const movements: IMovementItem[] = movementRecords.map((item) => {
        const material = materialMap.get(item.Material_Code) || materialMap.get(item.Material_Code || '');
        return {
        ID: item.ID,
        Title: item.Title || '',
        MovementType: item.Movement_Type || '',
        MaterialCode: item.Material_Code,
        MaterialName: material?.Material_Name || item.Material_Code,
        Category: material?.Category || '',
        QRCodeURL: material?.QRCodeURL || material?.qrcodeurl || '',
        Qty: item.Qty || 0,
        FromLocation: item.From_Location || '',
        ToLocation: item.To_Location || '',
        ProjectCode: item.Project_Code || '',
        Note: item.Note || '',
        Created: item.Created || '',
      };
      });
      setItems(movements);
    } catch (err) {
      console.error('Error fetching movements:', err);
      setError(err instanceof Error ? err.message : 'Failed to load movements');
    } finally {
      setIsLoading(false);
    }
  }, [sharePointService]);

  React.useEffect(() => {
    fetchMovements().catch(() => undefined);
  }, [fetchMovements]);

  React.useEffect(() => {
    const handleRefresh = (): void => {
      fetchMovements().catch(() => undefined);
    };
    window.addEventListener('refreshData', handleRefresh);
    return (): void => {
      window.removeEventListener('refreshData', handleRefresh);
    };
  }, [fetchMovements]);

  React.useEffect(() => {
    let filtered = [...items];

    if (selectedType !== 'all') {
      filtered = filtered.filter((item) => item.MovementType === selectedType);
    }

    if (startDate) {
      filtered = filtered.filter((item) => new Date(item.Created) >= startDate);
    }
    if (endDate) {
      const endOfDay = new Date(endDate);
      endOfDay.setHours(23, 59, 59, 999);
      filtered = filtered.filter((item) => new Date(item.Created) <= endOfDay);
    }

    if (searchText) {
      const lower = searchText.toLowerCase();
      filtered = filtered.filter(
        (item) =>
          item.MaterialCode.toLowerCase().includes(lower) ||
          item.MaterialName.toLowerCase().includes(lower) ||
          item.Category.toLowerCase().includes(lower) ||
          item.QRCodeURL.toLowerCase().includes(lower) ||
          item.ProjectCode.toLowerCase().includes(lower) ||
          item.FromLocation.toLowerCase().includes(lower) ||
          item.ToLocation.toLowerCase().includes(lower) ||
          item.Note.toLowerCase().includes(lower)
      );
    }

    setFilteredItems(filtered);
  }, [selectedType, startDate, endDate, searchText, items]);

  const countByType = (type: string): number =>
    filteredItems.filter((i) => i.MovementType === type).length;

  const totalQty = filteredItems.reduce((sum, item) => sum + item.Qty, 0);

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'exportCsv',
      text: 'Export CSV',
      iconProps: { iconName: 'ExcelDocument' },
      disabled: filteredItems.length === 0,
      onClick: () => exportMovementsToCsv(filteredItems),
    },
    {
      key: 'qrScan',
      text: 'QR Scan & Move',
      iconProps: { iconName: 'QRCode' },
      onClick: () => onQrScanRequested({ source: 'movements', movementType: 'Issue' }),
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        fetchMovements().catch(() => undefined);
        onRefresh();
      },
    },
    {
      key: 'clearFilters',
      text: 'Clear Filters',
      iconProps: { iconName: 'ClearFilter' },
      onClick: () => {
        setSelectedType('all');
        setSearchText('');
        setStartDate(undefined);
        setEndDate(undefined);
      },
    },
  ];

  if (isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ padding: '60px' }}>
        <Spinner size={SpinnerSize.large} label="Loading movements..." />
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

  return (
    <div className={classNames.container}>
      <div className={classNames.summaryRow}>
        <div className={classNames.summaryCard}>
          <Icon iconName="Switch" style={{ fontSize: 20, color: '#2563EB' }} />
          <div>
            <div className={classNames.cardValue}>{filteredItems.length}</div>
            <div className={classNames.cardLabel}>Movements</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="BoxAdditionSolid" style={{ fontSize: 20, color: '#059669' }} />
          <div>
            <div className={classNames.cardValue}>{countByType('GRN')}</div>
            <div className={classNames.cardLabel}>GRN</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="Forward" style={{ fontSize: 20, color: '#1E40AF' }} />
          <div>
            <div className={classNames.cardValue}>{countByType('Transfer Out')}</div>
            <div className={classNames.cardLabel}>Transfers</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="DeliveryTruck" style={{ fontSize: 20, color: '#92400E' }} />
          <div>
            <div className={classNames.cardValue}>{countByType('Issue')}</div>
            <div className={classNames.cardLabel}>Issues</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="Undo" style={{ fontSize: 20, color: '#3730A3' }} />
          <div>
            <div className={classNames.cardValue}>{countByType('Return')}</div>
            <div className={classNames.cardLabel}>Returns</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="Quantity" style={{ fontSize: 20, color: '#6B7280' }} />
          <div>
            <div className={classNames.cardValue}>{totalQty.toLocaleString()}</div>
            <div className={classNames.cardLabel}>Total Qty</div>
          </div>
        </div>
      </div>

      <div className={classNames.filterRow}>
        <SearchBox
          placeholder="Search material, project, location..."
          onChange={(_, value) => setSearchText(value || '')}
          styles={{ root: { width: isMobileView ? '100%' : '260px' } }}
        />
        <Dropdown
          placeholder="Movement type"
          options={TYPE_OPTIONS}
          selectedKey={selectedType}
          onChange={(_, option) => setSelectedType(option?.key as string || 'all')}
          styles={{ root: { width: isMobileView ? '100%' : '200px' } }}
        />
        <DatePicker
          placeholder="Start date"
          value={startDate}
          onSelectDate={(date) => setStartDate(date || undefined)}
          styles={{ root: { width: isMobileView ? '100%' : '160px' } }}
        />
        <DatePicker
          placeholder="End date"
          value={endDate}
          onSelectDate={(date) => setEndDate(date || undefined)}
          styles={{ root: { width: isMobileView ? '100%' : '160px' } }}
        />
        <CommandBar items={commandBarItems} styles={{ root: { padding: 0 } }} />
      </div>

      {filteredItems.length === 0 ? (
        <div className={classNames.emptyState}>
          <Icon iconName="Switch" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No movements found</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4, marginBottom: 12 }}>
            Movements are refreshed from the current inventory movement and material master registers.
          </Text>
          <Link underline onClick={() => onQrScanRequested({ source: 'movements', movementType: 'Issue' })}>
            Open QR Scan & Move
          </Link>
        </div>
      ) : (
        <EnhancedDataGrid
          listName={SHAREPOINT_LISTS.INVENTORY_MOVEMENTS_REGISTER}
          columns={[
            { key: 'created', name: 'Date', fieldName: 'Created', minWidth: 120,
              onRender: (item: IMovementItem) => item.Created ? new Date(item.Created).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—' },
            { key: 'movementType', name: 'Type', fieldName: 'MovementType', minWidth: 110, onRender: (item: IMovementItem) => <Tag text={item.MovementType} /> },
            { key: 'materialCode', name: 'Material Code', fieldName: 'MaterialCode', minWidth: 140,
              onRender: (item: IMovementItem) => <span style={{ fontWeight: 600, fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontSize: 13 }}>{item.MaterialCode}</span> },
            { key: 'materialName', name: 'Material Name', fieldName: 'MaterialName', minWidth: 180 },
            { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 60, onRender: (item: IMovementItem) => <span style={{ fontWeight: 600 }}>{item.Qty}</span> },
            { key: 'from', name: 'From', fieldName: 'FromLocation', minWidth: 100 },
            { key: 'to', name: 'To', fieldName: 'ToLocation', minWidth: 100 },
            { key: 'project', name: 'Project', fieldName: 'ProjectCode', minWidth: 120 },
          ]}
          spHttpClient={spHttpClient}
          pageContext={pageContext}
          showExport
          showSearch={false}
          showRefresh={false}
          showNewRecord={false}
        />
      )}
    </div>
  );
};

export default MovementsTab;
