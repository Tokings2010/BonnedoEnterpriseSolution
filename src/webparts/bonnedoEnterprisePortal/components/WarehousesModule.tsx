import * as React from 'react';
import {
  getTheme,
  mergeStyleSets,
  PrimaryButton,
  IconButton,
  Text,
  Icon,
  SearchBox,
  IDropdownOption,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import EnhancedDataGrid, { IDataGridColumn } from './EnhancedDataGrid';
import WarehouseForm from './WarehouseForm';
import WarehouseDetailsPanel from './WarehouseDetailsPanel';
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';
import { Tag, ProgressBar } from './TagRenderer';

export interface IWarehousesModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
}

export interface IWarehouse {
  ID: number;
  WarehouseCode: string;
  Warehouse_Name: string;
  Location: string;
  Status: string;
  [key: string]: any;
}

interface IBinInfo {
  id: string;
  occupied: boolean;
  full: boolean;
}

const generateBinGrid = (warehouseCode: string): IBinInfo[] => {
  const bins: IBinInfo[] = [];
  const occupancy = {
    'WH-PHC': 0.57, 'WH-ABJ': 0.34, 'WH-UBE': 0.82, 'WH-AKK': 0.45, 'WH-LOS': 0.21,
  };
  const occ = occupancy[warehouseCode as keyof typeof occupancy] || 0.4;
  for (let row = 1; row <= 3; row++) {
    for (let col = 1; col <= 8; col++) {
      const idx = (row - 1) * 8 + col;
      const r = Math.random();
      bins.push({
        id: `${['A', 'B', 'C'][row - 1]}${col}`,
        occupied: r < occ + 0.15,
        full: r < occ - 0.1,
      });
    }
  }
  return bins;
};

const WarehousesModule: React.FC<IWarehousesModuleProps> = ({ spHttpClient, pageContext }) => {
  const theme = getTheme();
  const [selectedWarehouse, setSelectedWarehouse] = React.useState<IWarehouse | undefined>(undefined);
  const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
  const [isFormPanelOpen, setIsFormPanelOpen] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [statusFilter, setStatusFilter] = React.useState('');
  const [activeWarehouse, setActiveWarehouse] = React.useState('WH-PHC');
  const [binGrid, setBinGrid] = React.useState<IBinInfo[]>([]);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);

  React.useEffect(() => {
    const handleResize = (): void => setIsMobileView(window.innerWidth < 768);
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  React.useEffect(() => {
    setBinGrid(generateBinGrid(activeWarehouse));
  }, [activeWarehouse, refreshKey]);

  React.useEffect(() => {
    const handleGlobalRefresh = (): void => setRefreshKey(p => p + 1);
    window.addEventListener('refreshData', handleGlobalRefresh);
    return () => window.removeEventListener('refreshData', handleGlobalRefresh);
  }, []);

  const classNames = mergeStyleSets({
    root: { display: 'flex', flexDirection: 'column', height: '100%', backgroundColor: '#FFFFFF', gap: 20 },
    header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 },
    headerTitle: { flex: 1, minWidth: 200 },
    headerActions: { display: 'flex', gap: 8, flexWrap: 'wrap' },
    gridContainer: { flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' },
    twoCol: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 },
    card: { backgroundColor: '#FFFFFF', borderRadius: 8, boxShadow: '0 1px 3px rgba(0,0,0,0.04)', border: '1px solid #E8ECF0', overflow: 'hidden' },
    cardHead: { padding: '14px 18px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #F0F2F5' },
    cardHeadTitle: { fontSize: 14, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 8, color: '#2C3E50' },
    cardBody: { padding: '16px 18px' },
    binGrid: { display: 'grid', gridTemplateColumns: 'repeat(8, 1fr)', gap: 3 },
    bin: { aspectRatio: '1', border: '1px solid #E8ECF0', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 8, color: '#B0B8C4', cursor: 'pointer', backgroundColor: '#FFFFFF' },
    binOcc: { backgroundColor: 'rgba(0,184,148,0.08)', borderColor: '#00B894', color: '#00B894', fontWeight: 600 },
    binFull: { backgroundColor: 'rgba(243,156,18,0.08)', borderColor: '#F39C12', color: '#F39C12' },
    qrLabel: { border: '2px dashed #E8ECF0', borderRadius: 8, padding: '12px 14px', display: 'flex', alignItems: 'center', gap: 12, maxWidth: 280 },
    qrBox: { width: 56, height: 56, backgroundColor: '#F5F6FA', border: '1px solid #E8ECF0', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
    qrInfo: { fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontSize: 10 },
    qrCode: { fontWeight: 700, fontSize: 11, color: '#2C3E50' },
    qrDesc: { color: '#7F8C9B' },
    qrLoc: { color: '#B0B8C4', fontSize: 9 },
    summaryRow: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 14 },
    summaryCard: { padding: '16px 20px', backgroundColor: '#FFFFFF', borderRadius: 8, boxShadow: '0 1px 3px rgba(0,0,0,0.04)', border: '1px solid #E8ECF0', borderLeft: '4px solid #00B894' },
    summaryValue: { fontSize: 24, fontWeight: 700, color: '#2C3E50' },
    summaryLabel: { fontSize: 11, color: '#7F8C9B', marginTop: 1 },
  });

  const warehouseColumns: IDataGridColumn[] = [
    {
      key: 'WarehouseCode', name: 'Code', fieldName: 'WarehouseCode', minWidth: 100,
      onRender: (item: IWarehouse) => (
        <span style={{ fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontWeight: 600, fontSize: 13 }}>
          {item.WarehouseCode || item.Title}
        </span>
      ),
    },
    {
      key: 'Warehouse_Name', name: 'Warehouse Name', fieldName: 'Warehouse_Name', minWidth: 180,
      onRender: (item: IWarehouse) => item.Warehouse_Name || item.Title,
    },
    {
      key: 'Location', name: 'Location', fieldName: 'Location', minWidth: 150,
    },
    {
      key: 'Utilization', name: 'Utilization', fieldName: 'Status', minWidth: 120,
      onRender: () => {
        const pctMap: Record<string, number> = {
          'WH-PHC': 57, 'WH-ABJ': 34, 'WH-UBE': 82, 'WH-AKK': 45, 'WH-LOS': 21,
        };
        const pct = pctMap[activeWarehouse] || 40;
        return <ProgressBar percent={pct} color={pct > 70 ? '#F39C12' : '#00B894'} showLabel />;
      },
    },
    {
      key: 'Type', name: 'Type', fieldName: 'Status', minWidth: 100,
      onRender: (item: IWarehouse) => {
        const typeMap: Record<string, string> = {
          'WH-PHC': 'Main', 'WH-ABJ': 'Main', 'WH-UBE': 'Site', 'WH-AKK': 'Laydown', 'WH-LOS': 'Transit',
        };
        return <Tag text={typeMap[item.WarehouseCode] || 'Main'} />;
      },
    },
    {
      key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100,
      onRender: (item: IWarehouse) => <Tag text={item.Status || 'Active'} />,
    },
  ];

  const statusOptions: IDropdownOption[] = [
    { key: '', text: 'All Status' },
    { key: 'Active', text: 'Active' },
    { key: 'Inactive', text: 'Inactive' },
  ];

  const handleRefresh = (): void => setRefreshKey(p => p + 1);
  const handleRowSelected = (warehouse: IWarehouse): void => {
    const code = warehouse.WarehouseCode || warehouse.Title;
    setActiveWarehouse(code);
    setSelectedWarehouse({
      ID: warehouse.ID,
      WarehouseCode: code,
      Warehouse_Name: warehouse.Warehouse_Name || warehouse.Title,
      Location: warehouse.Location || '',
      Status: warehouse.Status || 'Active',
    });
    setIsDetailsPanelOpen(true);
  };

  const filterQuery = statusFilter ? `Status eq '${statusFilter}'` : undefined;

  return (
    <div className={classNames.root}>
      <div className={classNames.header}>
        <div className={classNames.headerTitle}>
          <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: 4, color: '#2C3E50' }}>
            Warehouses
          </Text>
          <Text variant="medium" block style={{ color: '#7F8C9B' }}>
            Multi-site warehouse management with bin mapping
          </Text>
        </div>
        <div className={classNames.headerActions}>
          <PrimaryButton text="+ Add Warehouse" onClick={() => setIsFormPanelOpen(true)}
            iconProps={{ iconName: 'Add' }}
            styles={{ root: { backgroundColor: '#00B894', border: 'none', borderRadius: 8 }, rootHovered: { backgroundColor: '#009E80' } }} />
          <IconButton iconProps={{ iconName: 'Refresh' }} onClick={handleRefresh} title="Refresh" />
        </div>
      </div>

      {/* KPI Row */}
      <div className={classNames.summaryRow}>
        <div className={classNames.summaryCard} style={{ borderLeftColor: '#00B894' }}>
          <div className={classNames.summaryValue}>5</div>
          <div className={classNames.summaryLabel}>Total Warehouses</div>
        </div>
        <div className={classNames.summaryCard} style={{ borderLeftColor: '#4A90D9' }}>
          <div className={classNames.summaryValue}>4</div>
          <div className={classNames.summaryLabel}>Active</div>
        </div>
        <div className={classNames.summaryCard} style={{ borderLeftColor: '#F39C12' }}>
          <div className={classNames.summaryValue}>82%</div>
          <div className={classNames.summaryLabel}>Max Utilization</div>
        </div>
        <div className={classNames.summaryCard} style={{ borderLeftColor: '#9B59B6' }}>
          <div className={classNames.summaryValue}>192</div>
          <div className={classNames.summaryLabel}>Total Bins</div>
        </div>
      </div>

      {/* Two column: DataGrid + Side Panel */}
      <div className={!isMobileView ? classNames.twoCol : ''} style={isMobileView ? { display: 'flex', flexDirection: 'column', gap: 16 } : {}}>
        <div className={classNames.gridContainer}>
          <EnhancedDataGrid
            key={`warehouses-${refreshKey}`}
            listName={SHAREPOINT_LISTS.WAREHOUSES_MASTER}
            columns={warehouseColumns}
            filterQuery={filterQuery}
            pageSize={20}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            onRowSelected={handleRowSelected}
            showExport
            title=""
            statusFilterOptions={statusOptions}
            statusFilterValue={statusFilter}
            onStatusFilterChange={setStatusFilter}
          />
        </div>

        {!isMobileView && (
          <div>
            {/* Bin Map */}
            <div className={classNames.card}>
              <div className={classNames.cardHead}>
                <div className={classNames.cardHeadTitle}>
                  <Icon iconName="MapPinSolid" style={{ color: '#00B894' }} />
                  {activeWarehouse} Bin Map
                </div>
                <SearchBox placeholder="Search bin..." styles={{ root: { width: 160, borderRadius: 6 } }} />
              </div>
              <div className={classNames.cardBody}>
                <div className={classNames.binGrid}>
                  {binGrid.map(bin => (
                    <div
                      key={bin.id}
                      className={`${classNames.bin} ${bin.occupied ? classNames.binOcc : ''} ${bin.full ? classNames.binFull : ''}`}
                      title={`Bin ${bin.id}${bin.full ? ' (Full)' : bin.occupied ? ' (Occupied)' : ''}`}
                    >
                      {bin.id}
                    </div>
                  ))}
                </div>
                <div style={{ display: 'flex', gap: 16, marginTop: 12, fontSize: 11, color: '#7F8C9B' }}>
                  <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, backgroundColor: 'rgba(0,184,148,0.08)', border: '1px solid #00B894', marginRight: 4, verticalAlign: 'middle' }} /> Occupied</span>
                  <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, backgroundColor: 'rgba(243,156,18,0.08)', border: '1px solid #F39C12', marginRight: 4, verticalAlign: 'middle' }} /> Full</span>
                  <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, backgroundColor: '#FFFFFF', border: '1px solid #E8ECF0', marginRight: 4, verticalAlign: 'middle' }} /> Empty</span>
                </div>
              </div>
            </div>

            {/* QR Label */}
            <div className={classNames.card} style={{ marginTop: 16 }}>
              <div className={classNames.cardHead}>
                <div className={classNames.cardHeadTitle}>
                  <Icon iconName="QRCode" style={{ color: '#9B59B6' }} />
                  QR Label
                </div>
              </div>
              <div className={classNames.cardBody}>
                <div className={classNames.qrLabel}>
                  <div className={classNames.qrBox}>
                    <Icon iconName="QRCode" style={{ fontSize: 26, color: '#B0B8C4' }} />
                  </div>
                  <div className={classNames.qrInfo}>
                    <div className={classNames.qrCode}>PIP-CSL-06-0001</div>
                    <div className={classNames.qrDesc}>6" CS Line Pipe</div>
                    <div className={classNames.qrLoc}>B2026-0034 | {activeWarehouse} | A3-R2-S4</div>
                  </div>
                </div>
                <div style={{ marginTop: 10, display: 'flex', gap: 6 }}>
                  <IconButton iconProps={{ iconName: 'Print' }} title="Print" />
                  <IconButton iconProps={{ iconName: 'Copy' }} title="Batch" />
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      <WarehouseForm
        isOpen={isFormPanelOpen}
        onDismiss={() => setIsFormPanelOpen(false)}
        onSubmitSuccess={() => { setIsFormPanelOpen(false); handleRefresh(); }}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
      />

      <WarehouseDetailsPanel
        isOpen={isDetailsPanelOpen}
        warehouse={selectedWarehouse}
        onDismiss={() => { setIsDetailsPanelOpen(false); setSelectedWarehouse(undefined); }}
        onRefresh={handleRefresh}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
      />
    </div>
  );
};

export default WarehousesModule;
