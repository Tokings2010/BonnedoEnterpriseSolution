import * as React from 'react';
import {
  Text, getTheme, mergeStyleSets, IconButton, Icon, Spinner, SpinnerSize,
  DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, ConstrainMode,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface ICashFlowForecastTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
}

interface IMonthlyForecast {
  month: string;
  monthLabel: string;
  expectedInflow: number;
  expectedOutflow: number;
  netFlow: number;
  cumulativeBalance: number;
}

const CashFlowForecastTab: React.FC<ICashFlowForecastTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
}) => {
  const theme = getTheme();
  const [loading, setLoading] = React.useState(true);
  const [forecast, setForecast] = React.useState<IMonthlyForecast[]>([]);
  const [totals, setTotals] = React.useState({ inflow: 0, outflow: 0, netFlow: 0, closingBalance: 0 });
  const [refreshKey, setRefreshKey] = React.useState(0);

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

  const loadForecast = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    try {
      const webUrl = pageContext.web.absoluteUrl;
      let totalInflow = 0, totalOutflow = 0;

      // Fetch approved payments (expected outflows)
      let payments: any[] = [];
      try {
        const pUrl = `${webUrl}/_api/web/lists/getByTitle('FIN_Payment_Request_Register')/items?$filter=Approval_Status eq 'Approved'&$select=Amount,Due_Date,Payment_Date&$top=500`;
        const pResp = await spHttpClient.get(pUrl, SPHttpClient.configurations.v1);
        if (pResp.ok) {
          const pData = await pResp.json();
          payments = pData.value || [];
        }
      } catch { /* silent */ }

      // Build monthly forecast (next 6 months)
      const now = new Date();
      const months: IMonthlyForecast[] = [];
      let cumulative = 0;

      for (let i = 0; i < 6; i++) {
        const dt = new Date(now.getFullYear(), now.getMonth() + i, 1);
        const monthKey = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`;
        const monthLabel = dt.toLocaleDateString('en-GB', { month: 'short', year: '2-digit' });

        const monthlyOutflow = payments
          .filter((p: any) => {
            const dueStr = p.Due_Date || p.Payment_Date || '';
            if (!dueStr) return false;
            const d = new Date(dueStr);
            return d.getFullYear() === dt.getFullYear() && d.getMonth() === dt.getMonth();
          })
          .reduce((s: number, p: any) => s + (p.Amount || 0), 0);

        const inflow = 0;
        const outflow = monthlyOutflow;
        const net = inflow - outflow;
        cumulative += net;

        months.push({
          month: monthKey,
          monthLabel,
          expectedInflow: inflow,
          expectedOutflow: outflow,
          netFlow: net,
          cumulativeBalance: cumulative,
        });

        totalInflow += inflow;
        totalOutflow += outflow;
      }

      setForecast(months);
      setTotals({
        inflow: totalInflow,
        outflow: totalOutflow,
        netFlow: totalInflow - totalOutflow,
        closingBalance: cumulative,
      });
    } catch { /* silent */ }
    setLoading(false);
  }, [spHttpClient, pageContext]);

  React.useEffect(() => { loadForecast().catch(() => undefined); }, [loadForecast, refreshKey]);

  const handleRefresh = (): void => { setRefreshKey(prev => prev + 1); onRefresh(); };

  const fmt = (n: number): string => {
    if (n >= 1000000) return `\u20A6${(n / 1000000).toFixed(1)}M`;
    if (n >= 1000) return `\u20A6${(n / 1000).toFixed(1)}K`;
    return `\u20A6${n.toFixed(0)}`;
  };

  const forecastColumns: IColumn[] = [
    {
      key: 'monthLabel', name: 'Month', fieldName: 'monthLabel', minWidth: 80, isResizable: true,
      onRender: (item?: IMonthlyForecast) => item ? <span style={{ fontWeight: 600 }}>{item.monthLabel}</span> : null,
    },
    {
      key: 'expectedInflow', name: 'Inflow', fieldName: 'expectedInflow', minWidth: 120, isResizable: true,
      onRender: (item?: IMonthlyForecast) => item ? <span style={{ color: '#107C10', fontWeight: 500 }}>{fmt(item.expectedInflow)}</span> : null,
    },
    {
      key: 'expectedOutflow', name: 'Outflow', fieldName: 'expectedOutflow', minWidth: 120, isResizable: true,
      onRender: (item?: IMonthlyForecast) => item ? <span style={{ color: '#D13438', fontWeight: 500 }}>{fmt(item.expectedOutflow)}</span> : null,
    },
    {
      key: 'netFlow', name: 'Net Flow', fieldName: 'netFlow', minWidth: 120, isResizable: true,
      onRender: (item?: IMonthlyForecast) => item ? (
        <span style={{
          fontWeight: 600, padding: '2px 8px', borderRadius: '10px', fontSize: '11px',
          backgroundColor: item.netFlow >= 0 ? '#DFF6DD' : '#FDE7E9',
          color: item.netFlow >= 0 ? '#107C10' : '#A80000',
        }}>
          {item.netFlow >= 0 ? '+' : ''}{fmt(item.netFlow)}
        </span>
      ) : null,
    },
    {
      key: 'cumulativeBalance', name: 'Cumulative', fieldName: 'cumulativeBalance', minWidth: 120, isResizable: true,
      onRender: (item?: IMonthlyForecast) => item ? (
        <span style={{ fontWeight: 700, color: item.cumulativeBalance >= 0 ? '#107C10' : '#D13438' }}>
          {fmt(item.cumulativeBalance)}
        </span>
      ) : null,
    },
  ];

  return (
    <div className={classNames.container}>
      <div className={classNames.header}>
        <div className={classNames.headerTitle}>
          <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
            Cash Flow Forecast
          </Text>
          <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
            6-month projected cash flow based on approved payments
          </Text>
        </div>
        <div className={classNames.headerActions}>
          <IconButton iconProps={{ iconName: 'Refresh' }} onClick={handleRefresh} title="Refresh" />
        </div>
      </div>

      {!loading && forecast.length > 0 && (
        <div className={classNames.summaryRow}>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Expected Inflow</div>
            <div className={classNames.summaryValue} style={{ color: '#107C10' }}>{fmt(totals.inflow)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Expected Outflow</div>
            <div className={classNames.summaryValue} style={{ color: '#D13438' }}>{fmt(totals.outflow)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Net Flow</div>
            <div className={classNames.summaryValue} style={{ color: totals.netFlow >= 0 ? '#107C10' : '#D13438' }}>{fmt(totals.netFlow)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Closing Balance</div>
            <div className={classNames.summaryValue} style={{ color: totals.closingBalance >= 0 ? '#107C10' : '#D13438' }}>{fmt(totals.closingBalance)}</div>
          </div>
        </div>
      )}

      {loading ? (
        <div className={classNames.loadingContainer}><Spinner size={SpinnerSize.medium} label="Loading forecast..." /></div>
      ) : forecast.length === 0 ? (
        <div className={classNames.emptyState}>
          <Icon iconName="Timeline" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No forecast data</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4 }}>No approved payments found for forecasting.</Text>
        </div>
      ) : (
        <div className={classNames.gridContainer}>
          <DetailsList
            items={forecast}
            columns={forecastColumns}
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

export default CashFlowForecastTab;
