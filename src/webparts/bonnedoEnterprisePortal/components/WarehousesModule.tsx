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
  const grid = [
    // Row A: first 4 occupied, next 2 full, last 2 empty
    { id: 'A1', occupied: true, full: false },
    { id: 'A2', occupied: true, full: false },
    { id: 'A3', occupied: true, full: false },
    { id: 'A4', occupied: true, full: false },
    { id: 'A5', occupied: false, full: true },
    { id: 'A6', occupied: false, full: true },
    { id: 'A7', occupied: false, full: false },
    { id: 'A8', occupied: false, full: false },
    // Row B
    { id: 'B1', occupied: true, full: false },
    { id: 'B2', occupied: false, full: true },
    { id: 'B3', occupied: true, full: false },
    { id: 'B4', occupied: true, full: false },
    { id: 'B5', occupied: false, full: true },
    { id: 'B6', occupied: false, full: false },
    { id: 'B7', occupied: true, full: false },
    { id: 'B8', occupied: false, full: false },
    // Row C
    { id: 'C1', occupied: true, full: false },
    { id: 'C2', occupied: true, full: false },
    { id: 'C3', occupied: false, full: true },
    { id: 'C4', occupied: false, full: false },
    { id: 'C5', occupied: true, full: false },
    { id: 'C6', occupied: true, full: false },
    { id: 'C7', occupied: false, full: true },
    { id: 'C8', occupied: false, full: false },
  ];
  // Shuffle based on warehouse code to give each warehouse a different layout
  const seed = warehouseCode.split('').reduce((s, c) => s + c.charCodeAt(0), 0);
  const shuffled = [...grid];
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = (seed + i * 7) % (i + 1);
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled;
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
    binOcc: { backgroundColor: 'rgba(0,120,212,0.08)', borderColor: '#0078D4', color: '#0078D4', fontWeight: 600 },
    binFull: { backgroundColor: 'rgba(243,156,18,0.08)', borderColor: '#F39C12', color: '#F39C12' },
    qrLabel: { border: '2px dashed #E8ECF0', borderRadius: 8, padding: '12px 14px', display: 'flex', alignItems: 'center', gap: 12, maxWidth: 280 },
    qrBox: { width: 56, height: 56, backgroundColor: '#F5F6FA', border: '1px solid #E8ECF0', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
    qrInfo: { fontFamily: "'Cascadia Code','Fira Code',Consolas,monospace", fontSize: 10 },
    qrCode: { fontWeight: 700, fontSize: 11, color: '#2C3E50' },
    qrDesc: { color: '#7F8C9B' },
    qrLoc: { color: '#B0B8C4', fontSize: 9 },
    summaryRow: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 14 },
    summaryCard: { padding: '16px 20px', backgroundColor: '#FFFFFF', borderRadius: 8, boxShadow: '0 1px 3px rgba(0,0,0,0.04)', border: '1px solid #E8ECF0', borderLeft: '4px solid #0078D4' },
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
      onRender: (item: IWarehouse) => {
        const pctMap: Record<string, number> = {
          'WH-PHC': 57, 'WH-ABJ': 34, 'WH-UBE': 82, 'WH-AKK': 45, 'WH-LOS': 21,
        };
        const pct = pctMap[item.WarehouseCode || item.Title] || 40;
        return <ProgressBar percent={pct} color={pct > 70 ? '#F39C12' : '#0078D4'} showLabel />;
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
            styles={{ root: { backgroundColor: '#0078D4', border: 'none', borderRadius: 8 }, rootHovered: { backgroundColor: '#106EBE' } }} />
          <IconButton iconProps={{ iconName: 'Refresh' }} onClick={handleRefresh} title="Refresh" />
        </div>
      </div>

      {/* KPI Row */}
      <div className={classNames.summaryRow}>
        <div className={classNames.summaryCard} style={{ borderLeftColor: '#0078D4' }}>
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
                  <Icon iconName="MapPinSolid" style={{ color: '#0078D4' }} />
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
                  <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, backgroundColor: 'rgba(0,120,212,0.08)', border: '1px solid #0078D4', marginRight: 4, verticalAlign: 'middle' }} /> Occupied</span>
                  <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, backgroundColor: 'rgba(243,156,18,0.08)', border: '1px solid #F39C12', marginRight: 4, verticalAlign: 'middle' }} /> Full</span>
                  <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, backgroundColor: '#FFFFFF', border: '1px solid #E8ECF0', marginRight: 4, verticalAlign: 'middle' }} /> Empty</span>
                </div>
              </div>
            </div>

            {/* QR Label */}
            <div className={classNames.card} style={{ marginTop: 16 }}>
              <div className={classNames.cardHead}>
                <div className={classNames.cardHeadTitle}>
                  <Icon iconName="QRCode" style={{ color: '#0078D4' }} />
                  Warehouse QR Label
                </div>
              </div>
              <div className={classNames.cardBody}>
                <div className={classNames.qrLabel}>
                  <div className={classNames.qrBox}>
                    <img
                      src={`https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=${encodeURIComponent(activeWarehouse)}`}
                      alt="QR"
                      style={{ width: 52, height: 52, objectFit: 'contain' }}
                    />
                  </div>
                  <div className={classNames.qrInfo}>
                    <div className={classNames.qrCode}>{activeWarehouse}</div>
                    <div className={classNames.qrDesc}>{selectedWarehouse?.Warehouse_Name || 'Warehouse'}</div>
                    <div className={classNames.qrLoc}>Scan to identify warehouse</div>
                  </div>
                </div>
                <div style={{ marginTop: 10, display: 'flex', gap: 6 }}>
                  <IconButton
                    iconProps={{ iconName: 'Print' }}
                    title="Print QR label"
                    onClick={() => {
                      const code = activeWarehouse;
                      const name = selectedWarehouse?.Warehouse_Name || 'Warehouse';
                      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=${encodeURIComponent(code)}`;
                      const w = window.open('', '_blank', 'width=400,height=500');
                      if (w) {
                        w.document.write(`
                          <html><head><title>QR Label - ${code}</title>
                          <style>
                            body { text-align: center; font-family: 'Segoe UI', sans-serif; padding: 40px 20px; }
                            .qr-image { width: 260px; height: 260px; margin: 20px auto; border: 1px solid #E1E1E1; padding: 16px; border-radius: 8px; }
                            .qr-image img { width: 100%; height: 100%; object-fit: contain; }
                            .code { font-size: 20px; font-weight: 700; font-family: 'Cascadia Code', monospace; margin: 12px 0; }
                            .name { font-size: 14px; color: #616161; }
                            @media print { @page { margin: 0; } body { padding: 20px; } }
                          </style></head>
                          <body>
                            <h2 style="font-size:20px;margin-bottom:4px;">Bonnedo Enterprise</h2>
                            <p style="color:#8A8A8A;margin-top:0;">Warehouse QR Label</p>
                            <div class="qr-image"><img src="${qrUrl}" alt="QR Code" /></div>
                            <div class="code">${code}</div>
                            <div class="name">${name}</div>
                          </body></html>
                        `);
                        w.document.close();
                      }
                    }}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Copy' }}
                    title="Copy warehouse code"
                    onClick={() => {
                      navigator.clipboard.writeText(activeWarehouse).catch(() => undefined);
                    }}
                  />
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
