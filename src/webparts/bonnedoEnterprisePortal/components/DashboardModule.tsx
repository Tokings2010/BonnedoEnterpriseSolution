import * as React from 'react';
import {
  Text,
  Dropdown,
  IDropdownOption,
  getTheme,
  mergeStyleSets,
  IconButton,
  Shimmer,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';
import DashboardBarChart from './charts/DashboardBarChart';
import DashboardDonutChart from './charts/DashboardDonutChart';
import ProgressDashboardCard from './ProgressDashboardCard';

export interface IDashboardModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  onNavigate?: (module: string) => void;
}

interface IDashboardMetrics {
  totalProjects: number; activeProjects: number; completedProjects: number;
  totalMaterialRequests: number; pendingMaterialRequests: number;
  totalPurchaseRequisitions: number; pendingPurchaseRequisitions: number;
  totalPurchaseOrders: number; pendingPurchaseOrders: number; totalGoodsReceived: number;
  totalPaymentRequests: number; pendingPaymentRequests: number;
  totalApprovedPayments: number; totalExpenses: number; totalPaymentAmount: number;
  totalMaterials: number; totalVendors: number;
  totalInventoryItems: number; lowStockItems: number; totalMovements: number;
}

interface IActivityItem {
  id: number; title: string; description: string; timestamp: string;
  type: 'project' | 'procurement' | 'finance' | 'material' | 'inventory';
}

function computeLowStockCount(inventory: any[], materials: any[]): number {
  const minMap = new Map<string, number>();
  materials.forEach(m => { if (m.Material_Code && m.MinStockLevel > 0) minMap.set(String(m.Material_Code), Number(m.MinStockLevel)); });
  return inventory.filter(inv => {
    const code = inv.Material_Code || inv.field_1;
    const minLevel = minMap.get(String(code));
    return minLevel !== undefined && (inv.Qty_On_Hand || 0) < minLevel;
  }).length;
}

const DashboardModule: React.FC<IDashboardModuleProps> = ({ spHttpClient, pageContext, userDisplayName, onNavigate }) => {
  const theme = getTheme();
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [metrics, setMetrics] = React.useState<IDashboardMetrics>({
    totalProjects: 0, activeProjects: 0, completedProjects: 0,
    totalMaterialRequests: 0, pendingMaterialRequests: 0,
    totalPurchaseRequisitions: 0, pendingPurchaseRequisitions: 0,
    totalPurchaseOrders: 0, pendingPurchaseOrders: 0, totalGoodsReceived: 0,
    totalPaymentRequests: 0, pendingPaymentRequests: 0, totalApprovedPayments: 0,
    totalExpenses: 0, totalPaymentAmount: 0, totalMaterials: 0, totalVendors: 0,
    totalInventoryItems: 0, lowStockItems: 0, totalMovements: 0,
  });
  const [recentActivities, setRecentActivities] = React.useState<IActivityItem[]>([]);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [selectedProjectCode, setSelectedProjectCode] = React.useState<string>('');
  const [projectOptions, setProjectOptions] = React.useState<IDropdownOption[]>([]);

  const loadDashboardData = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(null);
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const [projects, materialRequests, purchaseRequisitions, purchaseOrders, goodsReceivedNotes, paymentRequests, expenses, materials, vendors, inventory, movements] =
        await Promise.all([
          fetchListData(webUrl, 'ENT_Project_Master'),
          fetchListData(webUrl, 'PRC_Material_Request_Register'),
          fetchListData(webUrl, 'PRC_Purchase_Requisition_Register'),
          fetchListData(webUrl, 'PRC_Purchase_Order_Register'),
          fetchListData(webUrl, 'PRC_GRN_Register'),
          fetchListData(webUrl, 'FIN_Payment_Request_Register'),
          fetchListData(webUrl, 'FIN_Expense_Register'),
          fetchListData(webUrl, 'ENT_Materials_Master'),
          fetchListData(webUrl, 'ENT_Vendors_Master'),
          fetchListData(webUrl, SHAREPOINT_LISTS.INVENTORY_REGISTER),
          fetchListData(webUrl, SHAREPOINT_LISTS.INVENTORY_MOVEMENTS_REGISTER),
        ]);
      const active = projects.filter((i: any) => i.Project_Status === 'Active' || i.Status === 'Active' || i.Status === 'In Progress').length;
      const completed = projects.filter((i: any) => i.Project_Status === 'Completed' || i.Status === 'Completed').length;
      const pendMR = materialRequests.filter((i: any) => i.Approval_Status === 'Pending' || i.Status === 'Pending').length;
      const pendPR = purchaseRequisitions.filter((i: any) => i.Approval_Status === 'Pending' || i.Status === 'Pending').length;
      const pendPO = purchaseOrders.filter((i: any) => i.Approval_Status === 'Pending' || i.Status === 'Pending').length;
      const pendPay = paymentRequests.filter((i: any) => i.Approval_Status === 'Pending').length;
      const apprPay = paymentRequests.filter((i: any) => i.Approval_Status === 'Approved').length;
      const payAmt = paymentRequests.filter((i: any) => i.Approval_Status === 'Approved').reduce((s: number, i: any) => s + (i.Amount || 0), 0);
      setMetrics({
        totalProjects: projects.length, activeProjects: active, completedProjects: completed,
        totalMaterialRequests: materialRequests.length, pendingMaterialRequests: pendMR,
        totalPurchaseRequisitions: purchaseRequisitions.length, pendingPurchaseRequisitions: pendPR,
        totalPurchaseOrders: purchaseOrders.length, pendingPurchaseOrders: pendPO,
        totalGoodsReceived: goodsReceivedNotes.length,
        totalPaymentRequests: paymentRequests.length, pendingPaymentRequests: pendPay,
        totalApprovedPayments: apprPay, totalExpenses: expenses.length, totalPaymentAmount: payAmt,
        totalMaterials: materials.length, totalVendors: vendors.length,
        totalInventoryItems: inventory.length, lowStockItems: computeLowStockCount(inventory, materials),
        totalMovements: movements.length,
      });

      // Populate project dropdown for ProgressDashboardCard
      setProjectOptions(
        projects.map((p: any) => ({
          key: p.Project_Code || p.Title || '',
          text: `${p.Project_Code || ''} — ${p.Project_Name || p.Title || ''}`,
        })).filter((opt: IDropdownOption) => opt.key)
      );
      const acts: IActivityItem[] = [];
      projects.slice(0, 2).forEach((i: any) => acts.push({ id: i.ID, title: `Project: ${i.Project_Name || i.Title || i.Project_Code || i.ID}`, description: `Status: ${i.Project_Status || i.Status || 'Active'}`, timestamp: i.Created || new Date().toISOString(), type: 'project' }));
      materialRequests.slice(0, 2).forEach((i: any) => acts.push({ id: i.ID + 1000, title: `Material Request: ${i.Title || i.ID}`, description: `Status: ${i.Approval_Status || i.Status || 'Pending'}`, timestamp: i.Created || new Date().toISOString(), type: 'material' }));
      purchaseOrders.slice(0, 2).forEach((i: any) => acts.push({ id: i.ID + 2000, title: `Purchase Order: ${i.Title || i.PO_Number || i.ID}`, description: `Status: ${i.Approval_Status || i.Status || 'Pending'}`, timestamp: i.Created || new Date().toISOString(), type: 'procurement' }));
      paymentRequests.slice(0, 2).forEach((i: any) => acts.push({ id: i.ID + 3000, title: `Payment: ${i.Payment_Number || i.ID}`, description: `Status: ${i.Approval_Status || 'Pending'}`, timestamp: i.Created || new Date().toISOString(), type: 'finance' }));
      movements.slice(0, 2).forEach((i: any) => acts.push({ id: i.ID + 4000, title: `Movement: ${i.Movement_Type || 'Update'} \u2014 ${i.Material_Code || i.field_2 || i.Title}`, description: `Qty: ${i.Qty || 0}${i.From_Location ? ' from ' + i.From_Location : ''}${i.To_Location || i.To_x0020_Location ? ' to ' + (i.To_Location || i.To_x0020_Location) : ''}`, timestamp: i.Created || new Date().toISOString(), type: 'inventory' }));
      acts.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
      setRecentActivities(acts.slice(0, 8));
      setIsLoading(false);
    } catch (err) { console.error('Dashboard error:', err); setError('Failed to load dashboard data.'); setIsLoading(false); }
  }, [pageContext, refreshKey]);

  React.useEffect(() => { loadDashboardData(); }, []);

  const fetchListData = async (webUrl: string, listName: string): Promise<any[]> => {
    try {
      const r = await spHttpClient.get(`${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=500`, SPHttpClient.configurations.v1);
      if (!r.ok) return [];
      const d = await r.json();
      return d.value || [];
    } catch { return []; }
  };

  const handleRefresh = (): void => { loadDashboardData(); };

  const fmtCur = (amount: number): string => {
    if (amount >= 1000000) return `N${(amount / 1000000).toFixed(1)}M`;
    if (amount >= 1000) return `N${(amount / 1000).toFixed(1)}K`;
    return `N${amount.toFixed(0)}`;
  };

  const fmtTime = (ts: string): string => {
    try {
      const d = new Date(ts), n = new Date(), m = Math.floor((n.getTime() - d.getTime()) / 60000);
      if (m < 1) return 'Just now'; if (m < 60) return `${m}m ago`;
      const h = Math.floor(m / 60); if (h < 24) return `${h}h ago`;
      const dy = Math.floor(h / 24); if (dy < 7) return `${dy}d ago`;
      return d.toLocaleDateString();
    } catch { return ''; }
  };

  const M = (v: number) => v === undefined || v === null ? 0 : v;

  // ── STYLES — Mirror Executive Dashboard's proven pattern ──
  const c = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      gap: '24px',
      fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    },
    headerBar: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'flex-start',
    },
    headerLeft: { display: 'flex', flexDirection: 'column', gap: '4px' },
    sectionTitle: {
      fontSize: '16px',
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
      marginBottom: '16px',
    },

    // ── STATS BAR ──────────────────────────────────────────
    statsBar: {
      display: 'flex',
      alignItems: 'center',
      gap: '24px',
      padding: '16px 24px',
      background: `linear-gradient(135deg, ${theme.palette.themePrimary}, ${theme.palette.themeDark})`,
      borderRadius: '12px',
      color: theme.palette.white,
      boxShadow: '0 4px 14px rgba(0,0,0,0.12)',
      flexWrap: 'wrap',
    },
    statsItem: {
      display: 'flex',
      alignItems: 'baseline',
      gap: '8px',
    },
    statsValue: {
      fontSize: '24px',
      fontWeight: 700,
      lineHeight: 1.3,
    },
    statsLabel: {
      fontSize: '12px',
      fontWeight: 600,
      opacity: 0.8,
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
    },
    statsDivider: {
      width: '1px',
      height: '32px',
      background: 'rgba(255,255,255,0.2)',
      flexShrink: 0,
    },

    // ── KPI GRID – 4 columns ──────────────────────────────
    // EXACT pattern from Executive Dashboard
    kpiGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(4, 1fr)',
      gap: '16px',
      '@media (max-width: 900px)': { gridTemplateColumns: 'repeat(2, 1fr)' },
      '@media (max-width: 500px)': { gridTemplateColumns: '1fr' },
    },
    kpiCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '12px',
      boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
      border: `1px solid ${theme.palette.neutralLighter}`,
      display: 'flex',
      flexDirection: 'column',
      gap: '8px',
      transition: 'transform 0.2s ease, box-shadow 0.2s ease',
      cursor: 'pointer',
      selectors: {
        ':hover': { transform: 'translateY(-2px)', boxShadow: '0 4px 12px rgba(0,0,0,0.08)' },
      },
    },
    kpiIconRow: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
    },
    kpiIconBox: {
      width: '44px',
      height: '44px',
      borderRadius: '12px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: '22px',
      flexShrink: 0,
    },
    kpiTrend: {
      fontSize: '12px',
      fontWeight: 600,
      padding: '2px 8px',
      borderRadius: '20px',
      lineHeight: 1.4,
      marginLeft: '8px',
      flexShrink: 0,
    },
    kpiValue: {
      fontSize: '28px',
      fontWeight: 700,
      color: theme.palette.neutralPrimary,
    },
    kpiLabel: {
      fontSize: '13px',
      fontWeight: 500,
      color: theme.palette.neutralSecondary,
    },
    kpiDescription: {
      fontSize: '11px',
      color: theme.palette.neutralTertiary,
    },

    // ── CHARTS GRID – 2 columns ──────────────────────────
    chartsRow: {
      display: 'grid',
      gridTemplateColumns: '1fr 1fr',
      gap: '20px',
      '@media (max-width: 1000px)': { gridTemplateColumns: '1fr' },
    },
    chartCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '12px',
      boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
      border: `1px solid ${theme.palette.neutralLighter}`,
      overflow: 'hidden',
    },
    chartTitle: {
      fontSize: '15px',
      fontWeight: 600,
      marginBottom: '16px',
      color: theme.palette.neutralPrimary,
    },

    // ── ACTIVITY ─────────────────────────────────────────
    // Larger fonts for readability
    activityCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '12px',
      boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
      border: `1px solid ${theme.palette.neutralLighter}`,
    },
    activityHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '16px',
    },
    activityCount: {
      fontSize: '13px',
      color: theme.palette.neutralTertiary,
    },
    activityList: {
      display: 'flex',
      flexDirection: 'column',
      gap: '8px',
      maxHeight: '340px',
      overflowY: 'auto',
    },
    activityItem: {
      display: 'flex',
      alignItems: 'center',
      gap: '14px',
      padding: '12px 14px',
      borderRadius: '8px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderLeft: '3px solid transparent',
      transition: 'background 0.12s ease',
      cursor: 'pointer',
      selectors: { ':hover': { backgroundColor: theme.palette.neutralLight } },
    },
    activityIconBox: {
      width: '40px',
      height: '40px',
      borderRadius: '8px',
      backgroundColor: theme.palette.white,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: '20px',
      flexShrink: 0,
    },
    activityContent: { flex: 1, minWidth: 0 },
    activityTopRow: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '4px',
    },
    activityTitle: {
      fontSize: '14px',
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
      whiteSpace: 'nowrap',
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      flex: 1,
      minWidth: 0,
    },
    activityBadge: {
      fontSize: '11px',
      fontWeight: 600,
      padding: '2px 8px',
      borderRadius: '10px',
      lineHeight: 1.4,
      flexShrink: 0,
      marginLeft: '8px',
    },
    activityBottomRow: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
    },
    activityProject: {
      fontSize: '13px',
      color: theme.palette.neutralSecondary,
      whiteSpace: 'nowrap',
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      flex: 1,
      minWidth: 0,
    },
    activityTime: {
      fontSize: '12px',
      color: theme.palette.neutralTertiary,
      flexShrink: 0,
      marginLeft: '8px',
    },

    // ── QUICK ACTIONS ────────────────────────────────────
    quickActionsRow: {
      display: 'flex',
      flexWrap: 'wrap',
      gap: '10px',
    },
    quickActionBtn: {
      display: 'inline-flex',
      alignItems: 'center',
      gap: '8px',
      padding: '10px 18px',
      backgroundColor: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLighter}`,
      borderRadius: '10px',
      cursor: 'pointer',
      fontSize: '14px',
      fontWeight: 500,
      color: theme.palette.neutralPrimary,
      transition: 'all 0.15s ease',
      boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
      selectors: {
        ':hover': {
          backgroundColor: theme.palette.neutralLighterAlt,
          borderColor: theme.palette.neutralLight,
          boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
        },
      },
    },

    // ── EMPTY STATE ──────────────────────────────────────
    emptyState: {
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '40px',
      textAlign: 'center',
    },
    emptyIcon: { fontSize: '36px', marginBottom: '12px', opacity: 0.4 },
    emptyTitle: { fontSize: '15px', fontWeight: 600, color: theme.palette.neutralSecondary, margin: 0 },
    emptySub: { fontSize: '13px', color: theme.palette.neutralTertiary, marginTop: '6px' },

    // ── SHIMMER ────────────────────────────────────────────
    shCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '12px',
      boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
      border: `1px solid ${theme.palette.neutralLighter}`,
    },
  });

  const getStatusColor = (status: string): string => {
    const s = (status || '').toLowerCase();
    if (s.includes('approved') || s.includes('completed')) return theme.palette.green;
    if (s.includes('pending')) return theme.palette.orange;
    if (s.includes('rejected') || s.includes('cancelled')) return theme.palette.red;
    return theme.palette.neutralSecondary;
  };

  const getActivityIcon = (type: string): string => {
    switch (type) {
      case 'project': return '\uD83D\uDCC1';
      case 'procurement': return '\uD83D\uDED2';
      case 'finance': return '\uD83D\uDCB0';
      case 'material': return '\uD83D\uDCE6';
      case 'inventory': return '\uD83D\uDD04';
      default: return '\uD83D\uDCCC';
    }
  };

  return (
    <div className={c.root}>
      {/* ═══ HEADER ═══ */}
      <div className={c.headerBar}>
        <div className={c.headerLeft}>
          <Text variant="xxLarge" style={{ fontSize: '26px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
            Operational Dashboard
          </Text>
          <Text variant="medium" style={{ fontSize: '14px', color: theme.palette.neutralSecondary }}>
            Welcome back, {userDisplayName} &mdash; Here's your enterprise overview
          </Text>
        </div>
        <IconButton
          iconProps={{ iconName: 'Refresh' }}
          onClick={handleRefresh}
          title="Refresh data"
          ariaLabel="Refresh data"
          styles={{ root: { backgroundColor: theme.palette.white, borderRadius: 8, boxShadow: '0 1px 3px rgba(0,0,0,0.06)', border: `1px solid ${theme.palette.neutralLighter}` } }}
        />
      </div>

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      {/* ═══ PROJECT STATS BAR ═══ */}
      <div className={c.statsBar}>
        <div className={c.statsItem} style={{ gap: '10px' }}>
          <span style={{ fontSize: '20px', lineHeight: 1 }}>&#128202;</span>
          <span className={c.statsLabel}>Projects</span>
        </div>
        <div className={c.statsDivider} />
        <div className={c.statsItem}>
          <span className={c.statsValue}>{M(metrics.totalProjects)}</span>
          <span className={c.statsLabel}>Total</span>
        </div>
        <div className={c.statsDivider} />
        <div className={c.statsItem}>
          <span className={c.statsValue}>{M(metrics.activeProjects)}</span>
          <span className={c.statsLabel}>Active</span>
        </div>
        <div className={c.statsDivider} />
        <div className={c.statsItem}>
          <span className={c.statsValue}>{M(metrics.completedProjects)}</span>
          <span className={c.statsLabel}>Completed</span>
        </div>
      </div>

      {/* ═══ KPI CARDS ═══ */}
      {isLoading ? (
        <div className={c.kpiGrid}>
          {[1, 2, 3, 4].map(i => (
            <div key={i} className={c.shCard}>
              <Shimmer width={44} height={44} />
              <div style={{ marginTop: 8 }}><Shimmer width="60%" height="32px" /></div>
              <Shimmer width="80%" height="14px" />
              <Shimmer width="40%" height="12px" />
            </div>
          ))}
        </div>
      ) : (
        <>
          {/* Row 1 */}
          <div>
            <div className={c.sectionTitle}>Key Metrics</div>
            <div className={c.kpiGrid}>
              {/* Active Projects */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('projects')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.blue}18`, color: theme.palette.blue }}>&#128193;</div>
                  <div className={c.kpiTrend} style={{ backgroundColor: `${theme.palette.blue}12`, color: theme.palette.blue }}>{M(metrics.totalProjects)} total</div>
                </div>
                <div className={c.kpiValue}>{M(metrics.activeProjects)}</div>
                <div className={c.kpiLabel}>Active Projects</div>
                <div className={c.kpiDescription}>Projects currently in progress</div>
              </div>
              {/* Pending Material Requests */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('procurement')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.orange}18`, color: theme.palette.orange }}>&#128230;</div>
                  <div className={c.kpiTrend} style={{
                    backgroundColor: M(metrics.pendingMaterialRequests) > 0 ? `${theme.palette.orange}12` : `${theme.palette.green}12`,
                    color: M(metrics.pendingMaterialRequests) > 0 ? theme.palette.orange : theme.palette.green,
                  }}>
                    {M(metrics.pendingMaterialRequests) > 0 ? `${M(metrics.pendingMaterialRequests)} pending` : 'All clear'}
                  </div>
                </div>
                <div className={c.kpiValue}>{M(metrics.pendingMaterialRequests)}</div>
                <div className={c.kpiLabel}>Pending Material Requests</div>
                <div className={c.kpiDescription}>Awaiting approval or processing</div>
              </div>
              {/* Pending Purchase Orders */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('procurement')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.purple}18`, color: theme.palette.purple }}>&#128722;</div>
                  <div className={c.kpiTrend} style={{ backgroundColor: `${theme.palette.purple}12`, color: theme.palette.purple }}>{M(metrics.totalPurchaseOrders)} total</div>
                </div>
                <div className={c.kpiValue}>{M(metrics.pendingPurchaseOrders)}</div>
                <div className={c.kpiLabel}>Pending Purchase Orders</div>
                <div className={c.kpiDescription}>Orders awaiting approval</div>
              </div>
              {/* Approved Payments */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('finance')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.green}18`, color: theme.palette.green }}>&#128179;</div>
                  <div className={c.kpiTrend} style={{ backgroundColor: `${theme.palette.green}12`, color: theme.palette.green }}>{M(metrics.totalApprovedPayments)} payments</div>
                </div>
                <div className={c.kpiValue}>{fmtCur(M(metrics.totalPaymentAmount))}</div>
                <div className={c.kpiLabel}>Approved Payments</div>
                <div className={c.kpiDescription}>Total payment value approved</div>
              </div>
              {/* Pending Payments */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('finance')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.orange}18`, color: theme.palette.orange }}>&#9203;</div>
                  <div className={c.kpiTrend} style={{
                    backgroundColor: M(metrics.pendingPaymentRequests) > 0 ? `${theme.palette.orange}12` : `${theme.palette.green}12`,
                    color: M(metrics.pendingPaymentRequests) > 0 ? theme.palette.orange : theme.palette.green,
                  }}>
                    {M(metrics.pendingPaymentRequests) > 0 ? `${M(metrics.pendingPaymentRequests)} pending` : 'All clear'}
                  </div>
                </div>
                <div className={c.kpiValue}>{M(metrics.pendingPaymentRequests)}</div>
                <div className={c.kpiLabel}>Pending Payments</div>
                <div className={c.kpiDescription}>Payment requests awaiting action</div>
              </div>
              {/* Materials Master */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('material')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.teal}18`, color: theme.palette.teal }}>&#128203;</div>
                  <div className={c.kpiTrend} style={{ backgroundColor: `${theme.palette.teal}12`, color: theme.palette.teal }}>{M(metrics.totalInventoryItems)} in stock</div>
                </div>
                <div className={c.kpiValue}>{M(metrics.totalMaterials)}</div>
                <div className={c.kpiLabel}>Materials Master</div>
                <div className={c.kpiDescription}>Registered material items</div>
              </div>
              {/* Low Stock Items */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('material')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.orange}18`, color: theme.palette.orange }}>&#9888;&#65039;</div>
                  <div className={c.kpiTrend} style={{
                    backgroundColor: M(metrics.lowStockItems) > 0 ? `${theme.palette.orange}12` : `${theme.palette.green}12`,
                    color: M(metrics.lowStockItems) > 0 ? theme.palette.orange : theme.palette.green,
                  }}>
                    {M(metrics.lowStockItems) > 0 ? `${M(metrics.lowStockItems)} need restock` : 'All stocked'}
                  </div>
                </div>
                <div className={c.kpiValue}>{M(metrics.lowStockItems)}</div>
                <div className={c.kpiLabel}>Low Stock Items</div>
                <div className={c.kpiDescription}>Items below minimum level</div>
              </div>
              {/* Registered Vendors */}
              <div className={c.kpiCard} onClick={() => onNavigate?.('material')}>
                <div className={c.kpiIconRow}>
                  <div className={c.kpiIconBox} style={{ backgroundColor: `${theme.palette.blue}18`, color: theme.palette.blue }}>&#127970;</div>
                </div>
                <div className={c.kpiValue}>{M(metrics.totalVendors)}</div>
                <div className={c.kpiLabel}>Registered Vendors</div>
                <div className={c.kpiDescription}>Approved vendor partners</div>
              </div>
            </div>
          </div>

          {/* ═══ CHARTS ═══ */}
          <div>
            <div className={c.sectionTitle}>Analytics</div>
            <div className={c.chartsRow}>
              <div className={c.chartCard}>
                <div className={c.chartTitle}>Project Status Distribution</div>
                <DashboardDonutChart
                  data={[
                    { label: 'Active', value: M(metrics.activeProjects), color: theme.palette.blue },
                    { label: 'Completed', value: M(metrics.completedProjects), color: theme.palette.green },
                    { label: 'Other', value: Math.max(0, M(metrics.totalProjects) - M(metrics.activeProjects) - M(metrics.completedProjects)), color: theme.palette.neutralTertiary },
                  ]}
                  total={M(metrics.totalProjects)}
                />
              </div>
              <div className={c.chartCard}>
                <div className={c.chartTitle}>Procurement Pipeline</div>
                <DashboardBarChart
                  data={[
                    { label: 'MR', value: M(metrics.totalMaterialRequests), color: theme.palette.orange },
                    { label: 'PR', value: M(metrics.totalPurchaseRequisitions), color: theme.palette.yellow },
                    { label: 'PO', value: M(metrics.totalPurchaseOrders), color: theme.palette.purple },
                    { label: 'GRN', value: M(metrics.totalGoodsReceived), color: theme.palette.green },
                  ]}
                  height={220}
                />
              </div>
            </div>
            <div style={{ height: '20px', flexShrink: 0 }} />
            <div className={c.chartsRow}>
              <div className={c.chartCard}>
                <div className={c.chartTitle}>Finance Overview</div>
                <DashboardBarChart
                  data={[
                    { label: 'Approved', value: M(metrics.totalApprovedPayments), color: theme.palette.green },
                    { label: 'Pending', value: M(metrics.pendingPaymentRequests), color: theme.palette.orange },
                    { label: 'Expenses', value: M(metrics.totalExpenses), color: theme.palette.blue },
                  ]}
                  height={220}
                />
              </div>
              <div className={c.chartCard}>
                <div className={c.chartTitle}>Inventory Stock Status</div>
                <DashboardDonutChart
                  data={[
                    { label: 'In Stock', value: Math.max(0, M(metrics.totalInventoryItems) - M(metrics.lowStockItems)), color: theme.palette.green },
                    { label: 'Low Stock', value: M(metrics.lowStockItems), color: theme.palette.orange },
                  ]}
                  total={M(metrics.totalInventoryItems)}
                />
              </div>
            </div>
          </div>

          {/* ═══ PROJECT PROGRESS (ProgressDashboardCard) ═══ */}
          <div className={c.chartCard}>
            <div style={{ fontSize: '15px', fontWeight: 600, marginBottom: '12px', color: theme.palette.neutralPrimary, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span><span style={{ marginRight: 6 }}>&#128200;</span> Project Progress</span>
              {projectOptions.length > 0 && (
                <Dropdown
                  placeholder="Select a project..."
                  options={projectOptions}
                  selectedKey={selectedProjectCode}
                  onChange={(_, option) => setSelectedProjectCode(option?.key as string || '')}
                  styles={{ root: { minWidth: '260px' } }}
                />
              )}
            </div>
            <ProgressDashboardCard
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              projectCode={selectedProjectCode}
            />
          </div>

          {/* ═══ RECENT ACTIVITY ═══ */}
          <div className={c.activityCard}>
            <div className={c.activityHeader}>
              <div className={c.sectionTitle} style={{ marginBottom: 0 }}>Recent Activity</div>
              <span className={c.activityCount}>{recentActivities.length} updates</span>
            </div>
            {recentActivities.length > 0 ? (
              <div className={c.activityList}>
                {recentActivities.map(a => {
                  const status = a.description.replace('Status: ', '');
                  const sColor = getStatusColor(status);
                  return (
                    <div
                      key={a.id}
                      className={c.activityItem}
                      style={{ borderLeftColor: sColor }}
                      onClick={() => {
                        if (a.type === 'project') onNavigate?.('projects');
                        else if (a.type === 'procurement') onNavigate?.('procurement');
                        else if (a.type === 'finance') onNavigate?.('finance');
                        else if (a.type === 'material' || a.type === 'inventory') onNavigate?.('material');
                      }}
                    >
                      <div className={c.activityIconBox}>{getActivityIcon(a.type)}</div>
                      <div className={c.activityContent}>
                        <div className={c.activityTopRow}>
                          <span className={c.activityTitle}>{a.title}</span>
                          <span className={c.activityBadge} style={{ backgroundColor: `${sColor}18`, color: sColor }}>
                            {status}
                          </span>
                        </div>
                        <div className={c.activityBottomRow}>
                          <span className={c.activityProject}>{a.description}</span>
                          <span className={c.activityTime}>{fmtTime(a.timestamp)}</span>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className={c.emptyState}>
                <div className={c.emptyIcon}>&#128236;</div>
                <p className={c.emptyTitle}>No Recent Activity</p>
                <p className={c.emptySub}>Start creating projects and requests to see activity here</p>
              </div>
            )}
          </div>

          {/* ═══ QUICK ACTIONS ═══ */}
          <div className={c.chartCard}>
            <div className={c.chartTitle}>Quick Actions</div>
            <div className={c.quickActionsRow}>
              <button className={c.quickActionBtn} onClick={() => onNavigate?.('projects')}>
                <span style={{ fontSize: '16px', lineHeight: 1 }}>&#10133;</span> New Project
              </button>
              <button className={c.quickActionBtn} onClick={() => onNavigate?.('procurement')}>
                <span style={{ fontSize: '16px', lineHeight: 1 }}>&#128230;</span> Material Request
              </button>
              <button className={c.quickActionBtn} onClick={() => onNavigate?.('procurement')}>
                <span style={{ fontSize: '16px', lineHeight: 1 }}>&#128722;</span> Purchase Order
              </button>
              <button className={c.quickActionBtn} onClick={() => onNavigate?.('finance')}>
                <span style={{ fontSize: '16px', lineHeight: 1 }}>&#128179;</span> Payment Request
              </button>
              <button className={c.quickActionBtn} onClick={() => onNavigate?.('reports')}>
                <span style={{ fontSize: '16px', lineHeight: 1 }}>&#128200;</span> View Reports
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default DashboardModule;
