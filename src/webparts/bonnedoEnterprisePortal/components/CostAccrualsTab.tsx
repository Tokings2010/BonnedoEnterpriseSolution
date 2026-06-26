import * as React from 'react';
import {
  Text, getTheme, mergeStyleSets, IconButton, Icon, Spinner, SpinnerSize,
  DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, ConstrainMode,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface ICostAccrualsTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
}

interface IAccrualRecord {
  ID: number;
  Title: string;
  Project_Code: string;
  Vendor_Name: string;
  PO_Number: string;
  TotalAmount: number;
  Amount_Paid: number;
  Balance: number;
  Status: string;
  DeliveryDate: string;
  Transaction_Date: string;
}

/**
 * Aggregated cost transaction used for accrual rollup.
 */
interface ICostTransactionGroup {
  referenceId: string;
  projectCode: string;
  vendorName: string;
  commitmentAmount: number;
  paidAmount: number;
  txnDate: string;
}

const CostAccrualsTab: React.FC<ICostAccrualsTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
}) => {
  const theme = getTheme();
  const [loading, setLoading] = React.useState(true);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [records, setRecords] = React.useState<IAccrualRecord[]>([]);
  const [summary, setSummary] = React.useState({ totalPOValue: 0, totalPaid: 0, totalBalance: 0, count: 0 });

  const classNames = mergeStyleSets({
    container: { padding: '20px', height: '100%', display: 'flex', flexDirection: 'column' },
    header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' as const },
    headerTitle: { flex: 1, minWidth: '200px' },
    headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
    summaryRow: { display: 'flex', gap: '16px', flexWrap: 'wrap', marginBottom: '16px' },
    summaryCard: {
      flex: 1, minWidth: '160px', padding: '16px', backgroundColor: theme.palette.white,
      borderRadius: '8px', border: `1px solid ${theme.palette.neutralLight}`, boxShadow: theme.effects.elevation4,
    },
    summaryLabel: { fontSize: '12px', color: theme.palette.neutralSecondary, marginBottom: '4px' },
    summaryValue: { fontSize: '22px', fontWeight: 700, color: theme.palette.neutralPrimary },
    gridContainer: { flex: 1, overflow: 'auto' as const, minHeight: 0, backgroundColor: theme.palette.white, border: `1px solid ${theme.palette.neutralLight}`, borderRadius: '4px' },
    loadingContainer: { display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' },
    emptyState: { textAlign: 'center', padding: '60px 20px', color: theme.palette.neutralSecondary },
  });

  const loadData = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const groups = new Map<string, ICostTransactionGroup>();

      // ── PRIMARY: Fetch cost transactions from Cost_Transactions table ──
      try {
        const ctUrl = `${webUrl}/_api/web/lists/getByTitle('Cost_Transactions')/items?$top=2000&$orderby=Created desc`;
        const ctResp = await spHttpClient.get(ctUrl, SPHttpClient.configurations.v1);
        if (ctResp.ok) {
          const ctData = await ctResp.json();
          (ctData.value || []).forEach((tx: any) => {
            const refId: string = tx.Reference_ID || tx.Title || '';
            if (!refId) return;
            const amt = tx.Amount || 0;
            const txnType: string = (tx.Transaction_Type || '').toLowerCase();

            // Initialize group if new
            if (!groups.has(refId)) {
              groups.set(refId, {
                referenceId: refId,
                projectCode: tx.Project_Code || '',
                vendorName: tx.Vendor_Name || tx.Vendor || '',
                commitmentAmount: 0,
                paidAmount: 0,
                txnDate: tx.Transaction_Date || tx.Created || '',
              });
            }

            const group = groups.get(refId)!;
            // Update project/vendor from latest transaction
            if (tx.Project_Code) group.projectCode = tx.Project_Code;
            if (tx.Vendor_Name || tx.Vendor) group.vendorName = tx.Vendor_Name || tx.Vendor;
            if (tx.Transaction_Date) group.txnDate = tx.Transaction_Date;

            // Categorize by transaction type
            if (txnType.includes('po commitment') || txnType === 'purchase order') {
              group.commitmentAmount += amt;
            } else if (txnType.includes('vendor invoice') || txnType.includes('payment') || txnType.includes('paid')) {
              group.paidAmount += amt;
            } else if (txnType.includes('material issue') || txnType.includes('expense')) {
              group.paidAmount += amt;
            }
          });
        }
      } catch { /* silent */ }

      // ── FALLBACK: If no Cost_Transactions data, fall back to PO Register ──
      if (groups.size === 0) {
        try {
          const poUrl = `${webUrl}/_api/web/lists/getByTitle('PRC_Purchase_Order_Register')/items?$top=500&$orderby=Created desc`;
          const poResp = await spHttpClient.get(poUrl, SPHttpClient.configurations.v1);
          if (poResp.ok) {
            const poData = await poResp.json();
            (poData.value || []).forEach((po: any) => {
              const refId = po.PO_Number || po.Title || `PO-${po.ID}`;
              const totalAmt = po.TotalAmount || po.Amount || 0;
              const paidAmt = po.Amount_Paid || 0;
              if (totalAmt > 0) {
                groups.set(refId, {
                  referenceId: refId,
                  projectCode: po.Project_Code || po.Project_x0020_Code || '',
                  vendorName: typeof po.Vendor === 'object' ? (po.Vendor?.Title || '') : (po.Vendor || ''),
                  commitmentAmount: totalAmt,
                  paidAmount: paidAmt,
                  txnDate: po.Created || '',
                });
              }
            });
          }
        } catch { /* silent */ }
      }

      // ── Build records from grouped data ──
      const items: IAccrualRecord[] = [];
      let totalPO = 0, totalPaid = 0;

      groups.forEach((g) => {
        if (g.commitmentAmount <= 0) return;
        const balance = Math.max(0, g.commitmentAmount - g.paidAmount);
        items.push({
          ID: items.length + 1,
          Title: g.referenceId,
          Project_Code: g.projectCode,
          Vendor_Name: g.vendorName,
          PO_Number: g.referenceId,
          TotalAmount: g.commitmentAmount,
          Amount_Paid: g.paidAmount,
          Balance: balance,
          Status: balance <= 0 ? 'Fully Paid' : balance >= g.commitmentAmount ? 'Unpaid' : 'Partially Paid',
          DeliveryDate: '',
          Transaction_Date: g.txnDate,
        });
        totalPO += g.commitmentAmount;
        totalPaid += g.paidAmount;
      });

      // Sort by transaction date descending (newest first)
      items.sort((a, b) => b.Transaction_Date.localeCompare(a.Transaction_Date));

      setRecords(items);
      setSummary({
        totalPOValue: totalPO,
        totalPaid: totalPaid,
        totalBalance: totalPO - totalPaid,
        count: items.length,
      });
    } catch { /* silent */ }
    setLoading(false);
  }, [spHttpClient, pageContext]);

  React.useEffect(() => { loadData().catch(() => undefined); }, [loadData, refreshKey]);

  const handleRefresh = (): void => { setRefreshKey(prev => prev + 1); onRefresh(); };

  const fmt = (n: number): string => {
    if (n >= 1000000) return `\u20A6${(n / 1000000).toFixed(1)}M`;
    if (n >= 1000) return `\u20A6${(n / 1000).toFixed(1)}K`;
    return `\u20A6${n.toFixed(0)}`;
  };

  const accrualColumns: IColumn[] = [
    {
      key: 'PO_Number', name: 'PO Number', fieldName: 'PO_Number', minWidth: 120, isResizable: true,
      onRender: (item?: IAccrualRecord) => item ? <span style={{ fontWeight: 600, fontFamily: "'Cascadia Code', Consolas, monospace", fontSize: 12 }}>{item.PO_Number}</span> : null,
    },
    {
      key: 'Project_Code', name: 'Project', fieldName: 'Project_Code', minWidth: 100, isResizable: true,
    },
    {
      key: 'Vendor_Name', name: 'Vendor', fieldName: 'Vendor_Name', minWidth: 150, isResizable: true,
    },
    {
      key: 'TotalAmount', name: 'PO Value', fieldName: 'TotalAmount', minWidth: 130, isResizable: true,
      onRender: (item?: IAccrualRecord) => item ? <span style={{ fontWeight: 500 }}>{fmt(item.TotalAmount)}</span> : null,
    },
    {
      key: 'Amount_Paid', name: 'Paid', fieldName: 'Amount_Paid', minWidth: 110, isResizable: true,
      onRender: (item?: IAccrualRecord) => item ? <span style={{ color: '#107C10', fontWeight: 500 }}>{fmt(item.Amount_Paid)}</span> : null,
    },
    {
      key: 'Balance', name: 'Accrued Balance', fieldName: 'Balance', minWidth: 130, isResizable: true,
      onRender: (item?: IAccrualRecord) => item ? (
        <span style={{ fontWeight: 700, color: item.Balance > 0 ? '#D13438' : '#107C10' }}>
          {fmt(item.Balance)}
        </span>
      ) : null,
    },
    {
      key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, isResizable: true,
    },
  ];

  return (
    <div className={classNames.container}>
      <div className={classNames.header}>
        <div className={classNames.headerTitle}>
          <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
            Cost Accruals
          </Text>
          <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
            Accrued costs from outstanding purchase order balances
          </Text>
        </div>
        <div className={classNames.headerActions}>
          <IconButton iconProps={{ iconName: 'Refresh' }} onClick={handleRefresh} title="Refresh" />
        </div>
      </div>

      {!loading && records.length > 0 && (
        <div className={classNames.summaryRow}>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Total PO Value</div>
            <div className={classNames.summaryValue}>{fmt(summary.totalPOValue)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Total Paid</div>
            <div className={classNames.summaryValue} style={{ color: '#107C10' }}>{fmt(summary.totalPaid)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Total Outstanding</div>
            <div className={classNames.summaryValue} style={{ color: '#D13438' }}>{fmt(summary.totalBalance)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Open POs</div>
            <div className={classNames.summaryValue}>{summary.count}</div>
          </div>
        </div>
      )}

      {loading ? (
        <div className={classNames.loadingContainer}><Spinner size={SpinnerSize.medium} label="Loading accrual data..." /></div>
      ) : records.length === 0 ? (
        <div className={classNames.emptyState}>
          <Icon iconName="AccountActivity" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No accrual records found</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4 }}>
            Create Purchase Orders to see accrual tracking here.
          </Text>
        </div>
      ) : (
        <div className={classNames.gridContainer}>
          <DetailsList
            items={records}
            columns={accrualColumns}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            constrainMode={ConstrainMode.horizontalConstrained}
            selectionMode={SelectionMode.none}
            isHeaderVisible={true}
            compact={true}
          />
        </div>
      )}
    </div>
  );
};

export default CostAccrualsTab;
