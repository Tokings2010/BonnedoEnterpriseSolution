import * as React from 'react';
import {
  mergeStyleSets,
  Spinner, SpinnerSize,
  MessageBar, MessageBarType,
  Stack, Text, Icon,
  SearchBox, Dropdown, IDropdownOption,
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

export interface IInventoryTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
  onQrScanRequested: (context?: IQRScanRequestContext) => void;
}

interface IInventoryItem {
  ID: number;
  Title: string;
  MaterialCode: string;
  MaterialName: string;
  Category: string;
  QRCodeURL: string;
  Location: string;
  ProjectCode: string;
  QtyOnHand: number;
  QtyReserved: number;
  LastMovementDate: string;
  BinLocation: string;
  BatchNumber: string;
  Condition: string;
  DateReceived: string;
  Status: string;
}

interface IWarehouseOption {
  code: string;
  name: string;
}

interface IInventoryStyles {
  container: string;
  summaryRow: string;
  summaryCard: string;
  cardValue: string;
  cardLabel: string;
  filterRow: string;
  conditionGood: string;
  conditionDamaged: string;
  statusAvailable: string;
  statusReserved: string;
  emptyState: string;
}

const getStyles = (): IInventoryStyles =>
  mergeStyleSets({
    container: {
      display: 'flex',
      flexDirection: 'column',
      gap: '16px',
    },
    summaryRow: {
      display: 'flex',
      gap: '16px',
      flexWrap: 'wrap',
    },
    summaryCard: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '16px 24px',
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
      minWidth: '160px',
    },
    cardValue: {
      fontSize: '22px',
      fontWeight: '700',
      color: '#1E2532',
    },
    cardLabel: {
      fontSize: '12px',
      color: '#5A6A85',
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
    },
    filterRow: {
      display: 'flex',
      gap: '12px',
      alignItems: 'center',
      flexWrap: 'wrap',
    },
    conditionGood: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
      backgroundColor: '#D1FAE5',
      color: '#065F46',
    },
    conditionDamaged: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
      backgroundColor: '#FEF3C7',
      color: '#92400E',
    },
    statusAvailable: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
      backgroundColor: '#DBEAFE',
      color: '#1E40AF',
    },
    statusReserved: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
      backgroundColor: '#FEE2E2',
      color: '#991B1B',
    },
    emptyState: {
      textAlign: 'center',
      padding: '60px 20px',
      color: '#5A6A85',
    },
  });

const InventoryTab: React.FC<IInventoryTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
  onQrScanRequested,
}) => {
  const [items, setItems] = React.useState<IInventoryItem[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IInventoryItem[]>([]);
  const [warehouses, setWarehouses] = React.useState<IWarehouseOption[]>([]);
  const [minStockLevels, setMinStockLevels] = React.useState<Record<string, number>>({});
  const [selectedWarehouse, setSelectedWarehouse] = React.useState<string>('all');
  const [searchText, setSearchText] = React.useState('');
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  const classNames = getStyles();

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const isLowStock = React.useCallback(
    (materialCode: string, qtyOnHand: number): boolean => {
      const minLevel = minStockLevels[materialCode];
      return minLevel !== undefined && minLevel > 0 && qtyOnHand < minLevel;
    },
    [minStockLevels]
  );

  const fetchData = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(null);
    try {
      const [inventoryRecords, warehouseRecords, materialRecords] = await Promise.all([
        sharePointService.getInventoryRecords(),
        sharePointService.getWarehouses().catch(() => []),
        sharePointService.getMaterialMasterRecords().catch((): IMaterialMasterRecord[] => []),
      ]);

      const stockMap: Record<string, number> = {};
      materialRecords.forEach((m) => {
        if (m.Material_Code && m.MinStockLevel) {
          stockMap[m.Material_Code] = m.MinStockLevel;
        }
      });
      setMinStockLevels(stockMap);

      const materialMap = new Map<string, IMaterialMasterRecord>();
      materialRecords.forEach((material) => {
        const code = material.Material_Code || material.Title || '';
        if (code) {
          materialMap.set(code, material);
          materialMap.set(material.Title, material);
        }
      });

      const inventory: IInventoryItem[] = inventoryRecords.map((item) => {
        const material = materialMap.get(item.Material_Code) || materialMap.get(item.Material_Code || '');
        return {
        ID: item.ID,
        Title: item.Title || '',
        MaterialCode: item.Material_Code,
        MaterialName: material?.Material_Name || item.Material_Code,
        Category: material?.Category || '',
        QRCodeURL: material?.QRCodeURL || material?.qrcodeurl || '',
        Location: item.Location || '',
        ProjectCode: item.Project_Code || '',
        QtyOnHand: item.Qty_On_Hand || 0,
        QtyReserved: item.QtyReserved || 0,
        LastMovementDate: item.Last_Movement_Date || '',
        BinLocation: item.BinLocation || '',
        BatchNumber: item.BatchNumber || '',
        Condition: item.Condition || 'Good',
        DateReceived: item.DateReceived || '',
        Status: item.Status || 'Available',
      };
      });

      setItems(inventory);
      setWarehouses(
        warehouseRecords.map((wh) => ({
          code: wh.WarehouseCode || wh.Title || '',
          name: wh.Warehouse_Name || wh.Title || '',
        }))
      );
    } catch (err) {
      console.error('Error fetching inventory:', err);
      setError(err instanceof Error ? err.message : 'Failed to load inventory');
    } finally {
      setIsLoading(false);
    }
  }, [sharePointService]);

  React.useEffect(() => {
    fetchData().catch(() => undefined);
  }, [fetchData]);

  React.useEffect(() => {
    const handleRefresh = (): void => {
      fetchData().catch(() => undefined);
    };
    window.addEventListener('refreshData', handleRefresh);
    return (): void => {
      window.removeEventListener('refreshData', handleRefresh);
    };
  }, [fetchData]);

  React.useEffect(() => {
    let result = items;

    if (selectedWarehouse !== 'all') {
      result = result.filter((item) => item.Location === selectedWarehouse);
    }

    if (searchText) {
      const lower = searchText.toLowerCase();
      result = result.filter(
        (item) =>
          item.MaterialCode.toLowerCase().includes(lower) ||
          item.MaterialName.toLowerCase().includes(lower) ||
          item.Category.toLowerCase().includes(lower) ||
          item.QRCodeURL.toLowerCase().includes(lower) ||
          item.Location.toLowerCase().includes(lower) ||
          item.BinLocation.toLowerCase().includes(lower) ||
          item.BatchNumber.toLowerCase().includes(lower)
      );
    }

    setFilteredItems(result);
  }, [selectedWarehouse, searchText, items]);

  const totalItems = filteredItems.length;
  const totalQty = filteredItems.reduce((sum, item) => sum + item.QtyOnHand, 0);
  const warehouseCount = new Set(items.map((i) => i.Location).filter(Boolean)).size;
  const lowStockCount = filteredItems.filter((item: any) => isLowStock(item.Material_Code || item.MaterialCode, item.Qty_On_Hand || item.QtyOnHand)).length;

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'qrScan',
      text: 'QR Scan & Move',
      iconProps: { iconName: 'QRCode' },
      onClick: () => onQrScanRequested({ source: 'inventory', movementType: 'GRN' }),
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        fetchData().catch(() => undefined);
        onRefresh();
      },
    },
  ];

  const warehouseOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Warehouses' },
    ...warehouses.map((wh) => ({
      key: wh.code,
      text: `${wh.code} — ${wh.name}`,
    })),
  ];

  if (isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ padding: '60px' }}>
        <Spinner size={SpinnerSize.large} label="Loading inventory..." />
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
          <Icon iconName="NumberedList" style={{ fontSize: 20, color: '#2563EB' }} />
          <div>
            <div className={classNames.cardValue}>{totalItems}</div>
            <div className={classNames.cardLabel}>Line Items</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="BoxMultipleSolid" style={{ fontSize: 20, color: '#059669' }} />
          <div>
            <div className={classNames.cardValue}>{totalQty.toLocaleString()}</div>
            <div className={classNames.cardLabel}>Total Qty</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="CityNext" style={{ fontSize: 20, color: '#7C3AED' }} />
          <div>
            <div className={classNames.cardValue}>{warehouseCount}</div>
            <div className={classNames.cardLabel}>Warehouses</div>
          </div>
        </div>
        <div className={classNames.summaryCard}>
          <Icon iconName="Warning" style={{ fontSize: 20, color: '#D97706' }} />
          <div>
            <div className={classNames.cardValue} style={{ color: lowStockCount > 0 ? '#D97706' : undefined }}>
              {lowStockCount}
            </div>
            <div className={classNames.cardLabel}>Low Stock</div>
          </div>
        </div>
      </div>

      <div className={classNames.filterRow}>
        <SearchBox
          placeholder="Search material, warehouse, bin..."
          onChange={(_, value) => setSearchText(value || '')}
          styles={{ root: { width: isMobileView ? '100%' : '260px' } }}
        />
        <Dropdown
          placeholder="Filter by warehouse"
          options={warehouseOptions}
          selectedKey={selectedWarehouse}
          onChange={(_, option) => setSelectedWarehouse(option?.key as string || 'all')}
          styles={{ root: { width: isMobileView ? '100%' : '240px' } }}
        />
        <CommandBar items={commandBarItems} styles={{ root: { padding: 0 } }} />
      </div>

      {filteredItems.length === 0 ? (
        <div className={classNames.emptyState}>
          <Icon iconName="BoxMultipleSolid" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No inventory records</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4, marginBottom: 12 }}>
            Inventory is refreshed from the current material master and inventory register.
          </Text>
          <Link underline onClick={() => onQrScanRequested({ source: 'inventory', movementType: 'GRN' })}>
            Open QR Scan & Move
          </Link>
        </div>
      ) : (
        <EnhancedDataGrid
          listName={SHAREPOINT_LISTS.INVENTORY_REGISTER}
          columns={[
            { key: 'materialCode', name: 'Material Code', fieldName: 'Material_Code', minWidth: 140,
              onRender: (item: any) => <span style={{ fontWeight: 600, fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontSize: 13 }}>{item.Material_Code || item.MaterialCode || '—'}</span> },
            { key: 'project', name: 'Project Code', fieldName: 'Project_Code', minWidth: 100,
              onRender: (item: any) => <span>{item.Project_Code || item.ProjectCode || '—'}</span> },
            { key: 'location', name: 'Warehouse', fieldName: 'Location', minWidth: 100 },
            { key: 'qtyOnHand', name: 'Qty On Hand', fieldName: 'Qty_On_Hand', minWidth: 100,
              onRender: (item: any) => {
                const code = item.Material_Code || item.MaterialCode;
                const qty = item.Qty_On_Hand || item.QtyOnHand || 0;
                const low = isLowStock(code, qty);
                return <span style={{ fontWeight: 600, color: low ? '#F39C12' : '#107C10' }}>{qty}{low ? ' ⚠' : ''}</span>;
              }
            },
            { key: 'qtyReserved', name: 'Reserved', fieldName: 'QtyReserved', minWidth: 80,
              onRender: (item: any) => <span>{item.QtyReserved || 0}</span> },
            { key: 'binLocation', name: 'Bin', fieldName: 'BinLocation', minWidth: 70,
              onRender: (item: any) => <span>{item.BinLocation || '—'}</span> },
            { key: 'batchNumber', name: 'Batch', fieldName: 'BatchNumber', minWidth: 80,
              onRender: (item: any) => <span>{item.BatchNumber || '—'}</span> },
            { key: 'condition', name: 'Condition', fieldName: 'Condition', minWidth: 80 },
            { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 80 },
            { key: 'dateReceived', name: 'Date Received', fieldName: 'DateReceived', minWidth: 110,
              onRender: (item: any) => item.DateReceived ? new Date(item.DateReceived).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—' },
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

export default InventoryTab;
