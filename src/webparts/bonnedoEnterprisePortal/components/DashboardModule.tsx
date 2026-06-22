import * as React from 'react';
import {
  Text,
  getTheme,
  mergeStyleSets,
  IconButton,
  MessageBar,
  MessageBarType,
  Shimmer,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';

// Chart components
import DashboardBarChart from './charts/DashboardBarChart';
import DashboardDonutChart from './charts/DashboardDonutChart';
import DashboardLineChart from './charts/DashboardLineChart';

export interface IDashboardModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  onNavigate?: (module: string) => void;
}

// Dashboard data interfaces
interface IDashboardMetrics {
  // Projects (Core)
  totalProjects: number;
  activeProjects: number;
  completedProjects: number;

  // Procurement
  totalMaterialRequests: number;
  pendingMaterialRequests: number;
  totalPurchaseRequisitions: number;
  pendingPurchaseRequisitions: number;
  totalPurchaseOrders: number;
  pendingPurchaseOrders: number;
  totalGoodsReceived: number;

  // Finance
  totalPaymentRequests: number;
  pendingPaymentRequests: number;
  totalApprovedPayments: number;
  totalExpenses: number;
  totalPaymentAmount: number;

  // Materials
  totalMaterials: number;
  totalVendors: number;

  // Material Management / Inventory
  totalInventoryItems: number;
  lowStockItems: number;
  totalMovements: number;
}

interface IActivityItem {
  id: number;
  title: string;
  description: string;
  timestamp: string;
  type: 'project' | 'procurement' | 'finance' | 'material' | 'inventory';
}

function computeLowStockCount(inventory: any[], materials: any[]): number {
  const minMap = new Map<string, number>();
  materials.forEach((m) => {
    if (m.Material_Code && m.MinStockLevel > 0) {
      minMap.set(String(m.Material_Code), Number(m.MinStockLevel));
    }
  });
  return inventory.filter((inv) => {
    const code = inv.Material_Code || inv.field_1;
    const minLevel = minMap.get(String(code));
    if (minLevel === undefined) {
      return false;
    }
    return (inv.Qty_On_Hand || 0) < minLevel;
  }).length;
}

const DashboardModule: React.FC<IDashboardModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  onNavigate,
}) => {
  const theme = getTheme();
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [metrics, setMetrics] = React.useState<IDashboardMetrics>({
    totalProjects: 0, activeProjects: 0, completedProjects: 0,
    totalMaterialRequests: 0, pendingMaterialRequests: 0,
    totalPurchaseRequisitions: 0, pendingPurchaseRequisitions: 0,
    totalPurchaseOrders: 0, pendingPurchaseOrders: 0, totalGoodsReceived: 0,
    totalPaymentRequests: 0, pendingPaymentRequests: 0, totalApprovedPayments: 0,
    totalExpenses: 0, totalPaymentAmount: 0,
    totalMaterials: 0, totalVendors: 0,
    totalInventoryItems: 0, lowStockItems: 0, totalMovements: 0,
  });
  const [recentActivities, setRecentActivities] = React.useState<IActivityItem[]>([]);
  const [refreshKey, setRefreshKey] = React.useState(0);

  // Manual data fetch function (avoids hook-order issues on mount)
  const loadDashboardData = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(null);

    try {
      const webUrl = pageContext.web.absoluteUrl;

      const [
        projects,
        materialRequests,
        purchaseRequisitions,
        purchaseOrders,
        goodsReceivedNotes,
        paymentRequests,
        expenses,
        materials,
        vendors,
        inventory,
        movements,
      ] = await Promise.all([
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

      const activeProjects = projects.filter((item: any) =>
        item.Project_Status === 'Active' || item.Status === 'Active' || item.Status === 'In Progress'
      ).length;
      const completedProjects = projects.filter((item: any) =>
        item.Project_Status === 'Completed' || item.Status === 'Completed'
      ).length;

      const pendingMR = materialRequests.filter((item: any) =>
        item.Approval_Status === 'Pending' || item.Status === 'Pending'
      ).length;
      const pendingPR = purchaseRequisitions.filter((item: any) =>
        item.Approval_Status === 'Pending' || item.Status === 'Pending'
      ).length;
      const pendingPO = purchaseOrders.filter((item: any) =>
        item.Approval_Status === 'Pending' || item.Status === 'Pending'
      ).length;

      const pendingPayments = paymentRequests.filter((item: any) =>
        item.Approval_Status === 'Pending'
      ).length;
      const approvedPayments = paymentRequests.filter((item: any) =>
        item.Approval_Status === 'Approved'
      ).length;
      const totalPaymentAmount = paymentRequests
        .filter((item: any) => item.Approval_Status === 'Approved')
        .reduce((sum: number, item: any) => sum + (item.Amount || 0), 0);

      setMetrics({
        totalProjects: projects.length,
        activeProjects,
        completedProjects,
        totalMaterialRequests: materialRequests.length,
        pendingMaterialRequests: pendingMR,
        totalPurchaseRequisitions: purchaseRequisitions.length,
        pendingPurchaseRequisitions: pendingPR,
        totalPurchaseOrders: purchaseOrders.length,
        pendingPurchaseOrders: pendingPO,
        totalGoodsReceived: goodsReceivedNotes.length,
        totalPaymentRequests: paymentRequests.length,
        pendingPaymentRequests: pendingPayments,
        totalApprovedPayments: approvedPayments,
        totalExpenses: expenses.length,
        totalPaymentAmount,
        totalMaterials: materials.length,
        totalVendors: vendors.length,
        totalInventoryItems: inventory.length,
        lowStockItems: computeLowStockCount(inventory, materials),
        totalMovements: movements.length,
      });

      const activities: IActivityItem[] = [];
      projects.slice(0, 2).forEach((item: any) => {
        activities.push({
          id: item.ID,
          title: `Project: ${item.Project_Name || item.Title || item.Project_Code || item.ID}`,
          description: `Status: ${item.Project_Status || item.Status || 'Active'}`,
          timestamp: item.Created || new Date().toISOString(),
          type: 'project',
        });
      });
      materialRequests.slice(0, 2).forEach((item: any) => {
        activities.push({
          id: item.ID + 1000,
          title: `Material Request: ${item.Title || item.ID}`,
          description: `Status: ${item.Approval_Status || item.Status || 'Pending'}`,
          timestamp: item.Created || new Date().toISOString(),
          type: 'material',
        });
      });
      purchaseOrders.slice(0, 2).forEach((item: any) => {
        activities.push({
          id: item.ID + 2000,
          title: `Purchase Order: ${item.Title || item.PO_Number || item.ID}`,
          description: `Status: ${item.Approval_Status || item.Status || 'Pending'}`,
          timestamp: item.Created || new Date().toISOString(),
          type: 'procurement',
        });
      });
      paymentRequests.slice(0, 2).forEach((item: any) => {
        activities.push({
          id: item.ID + 3000,
          title: `Payment: ${item.Payment_Number || item.ID}`,
          description: `Status: ${item.Approval_Status || 'Pending'}`,
          timestamp: item.Created || new Date().toISOString(),
          type: 'finance',
        });
      });
      movements.slice(0, 2).forEach((item: any) => {
        activities.push({
          id: item.ID + 4000,
          title: `Movement: ${item.Movement_Type || 'Update'} — ${item.Material_Code || item.field_2 || item.Title}`,
          description: `Qty: ${item.Qty || 0}${item.From_Location ? ` from ${item.From_Location}` : ''}${item.To_Location || item.To_x0020_Location ? ` to ${item.To_Location || item.To_x0020_Location}` : ''}`,
          timestamp: item.Created || new Date().toISOString(),
          type: 'inventory',
        });
      });

      activities.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
      setRecentActivities(activities.slice(0, 6));
      setIsLoading(false);
    } catch (err) {
      console.error('Error fetching dashboard data:', err);
      setError('Failed to load dashboard data. Please refresh the page.');
      setIsLoading(false);
    }
  }, [pageContext, refreshKey]);

  // Initial load only once (prevents repeated hook-order issues)
  React.useEffect(() => {
    loadDashboardData();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const fetchListData = async (webUrl: string, listName: string): Promise<any[]> => {
    try {
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=500`;
      const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        console.warn(`Failed to fetch ${listName}:`, response.statusText);
        return [];
      }

      const data = await response.json();
      return data.value || [];
    } catch (err) {
      console.warn(`Error fetching ${listName}:`, err);
      return [];
    }
  };

  const handleRefresh = (): void => {
    loadDashboardData();
  };

  const formatCurrency = (amount: number): string => {
    if (amount >= 1000000) {
      return `N${(amount / 1000000).toFixed(1)}M`;
    } else if (amount >= 1000) {
      return `N${(amount / 1000).toFixed(1)}K`;
    }
    return `N${amount.toFixed(0)}`;
  };

  const formatTimestamp = (timestamp: string): string => {
    try {
      const date = new Date(timestamp);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffMins = Math.floor(diffMs / 60000);
      const diffHours = Math.floor(diffMins / 60);
      const diffDays = Math.floor(diffHours / 24);

      if (diffMins < 1) return 'Just now';
      if (diffMins < 60) return `${diffMins}m ago`;
      if (diffHours < 24) return `${diffHours}h ago`;
      if (diffDays < 7) return `${diffDays}d ago`;

      return date.toLocaleDateString();
    } catch {
      return 'Unknown';
    }
  };

  const getActivityIcon = (type: string): string => {
    switch (type) {
      case 'project': return 'Folder';
      case 'procurement': return 'ShoppingCart';
      case 'finance': return 'Money';
      case 'material': return 'Package';
      case 'inventory': return 'Switch';
      default: return 'Document';
    }
  };

  const getActivityColor = (type: string): string => {
    switch (type) {
      case 'project': return theme.palette.blue;
      case 'procurement': return theme.palette.purple;
      case 'finance': return theme.palette.green;
      case 'material': return theme.palette.orange;
      case 'inventory': return theme.palette.teal;
      default: return theme.palette.neutralSecondary;
    }
  };

  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      backgroundColor: theme.palette.neutralLighterAlt,
      padding: '24px',
      gap: '24px',
    },
    headerSection: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
    },
    headerTitle: {
      display: 'flex',
      flexDirection: 'column',
    },
    welcomeText: {
      fontSize: '24px',
      fontWeight: 700,
      color: theme.palette.neutralPrimary,
    },
    subWelcomeText: {
      fontSize: '14px',
      color: theme.palette.neutralSecondary,
    },
    sectionTitle: {
      fontSize: '18px',
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
      marginBottom: '16px',
    },
    metricsGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
      gap: '16px',
    },
    metricCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '8px',
      boxShadow: theme.effects.elevation4,
      transition: 'transform 0.2s ease, box-shadow 0.2s ease',
      cursor: 'pointer',
    },
    metricIcon: {
      fontSize: '28px',
      marginBottom: '12px',
    },
    metricValue: {
      fontSize: '28px',
      fontWeight: 700,
      marginBottom: '4px',
    },
    metricLabel: {
      fontSize: '13px',
      color: theme.palette.neutralSecondary,
      fontWeight: 500,
    },
    metricChange: {
      fontSize: '12px',
      marginTop: '8px',
    },
    chartsGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))',
      gap: '20px',
    },
    chartCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '8px',
      boxShadow: theme.effects.elevation4,
      overflow: 'hidden',
    },
    chartTitle: {
      fontSize: '16px',
      fontWeight: 600,
      marginBottom: '16px',
      color: theme.palette.neutralPrimary,
    },
    activityCard: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '8px',
      boxShadow: theme.effects.elevation4,
    },
    activityItem: {
      display: 'flex',
      alignItems: 'center',
      padding: '12px',
      marginBottom: '8px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRadius: '6px',
      gap: '12px',
      cursor: 'pointer',
      transition: 'background-color 0.2s ease',
    },
    activityIcon: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      width: '40px',
      height: '40px',
      borderRadius: '8px',
      fontSize: '18px',
    },
    activityContent: {
      flex: 1,
    },
    quickActionsGrid: {
      display: 'flex',
      flexWrap: 'wrap',
      gap: '12px',
    },
    quickActionButton: {
      padding: '12px 20px',
      backgroundColor: theme.palette.white,
      borderRadius: '8px',
      cursor: 'pointer',
      border: `1px solid ${theme.palette.neutralLight}`,
      color: theme.palette.neutralPrimary,
      fontWeight: 500,
      fontSize: '14px',
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      transition: 'all 0.2s ease',
      boxShadow: theme.effects.elevation4,
    },
    loadingContainer: {
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      minHeight: '400px',
      backgroundColor: theme.palette.neutralLighterAlt,
    },
    emptyState: {
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '40px',
      textAlign: 'center',
    },
    emptyIcon: {
      fontSize: '48px',
      marginBottom: '16px',
      opacity: 0.5,
    },
    emptyTitle: {
      fontSize: '16px',
      fontWeight: 600,
      color: theme.palette.neutralSecondary,
      marginBottom: '8px',
    },
    emptySubtitle: {
      fontSize: '13px',
      color: theme.palette.neutralTertiary,
    },
    statusBadge: {
      display: 'inline-block',
      padding: '2px 8px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: 500,
    },
    heroCard: {
      padding: '24px',
      background: `linear-gradient(135deg, ${theme.palette.themePrimary} 0%, ${theme.palette.themeDark} 100%)`,
      borderRadius: '12px',
      color: theme.palette.white,
    },
    heroTitle: {
      fontSize: '20px',
      fontWeight: 700,
      marginBottom: '8px',
    },
    heroSubtitle: {
      fontSize: '14px',
      opacity: 0.9,
      marginBottom: '16px',
    },
    heroStats: {
      display: 'flex',
      gap: '24px',
    },
    heroStat: {
      display: 'flex',
      flexDirection: 'column',
    },
    heroStatValue: {
      fontSize: '24px',
      fontWeight: 700,
    },
    heroStatLabel: {
      fontSize: '12px',
      opacity: 0.8,
    },
    summaryCard: {
      padding: '16px',
      backgroundColor: theme.palette.white,
      borderRadius: '8px',
      boxShadow: theme.effects.elevation4,
    },
  });

  return (
    <div className={classNames.root}>
      {/* Loading state with shimmer */}
      {isLoading ? (
        <>
          <div className={classNames.headerSection}>
            <div>
              <Shimmer width={200} height={32} />
              <Shimmer width={300} height={16} style={{ marginTop: 8 }} />
            </div>
          </div>

          <div className={classNames.metricsGrid}>
            {[1, 2, 3, 4, 5, 6].map((i) => (
              <div key={i} className={classNames.metricCard}>
                <Shimmer width={40} height={40} />
                <Shimmer width={80} height={28} style={{ marginTop: 12 }} />
                <Shimmer width={120} height={14} style={{ marginTop: 8 }} />
              </div>
            ))}
          </div>
        </>
      ) : (
        <>
          {/* Header */}
          <div className={classNames.headerSection}>
            <div className={classNames.headerTitle}>
              <Text variant="xxLarge" className={classNames.welcomeText}>
                Welcome back, {userDisplayName.split(' ')[0]}!
              </Text>
              <Text variant="medium" className={classNames.subWelcomeText}>
                Here is what is happening in your enterprise today.
              </Text>
            </div>
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              onClick={handleRefresh}
              title="Refresh data"
              ariaLabel="Refresh data"
            />
          </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          {error}
        </MessageBar>
      )}

      {/* Hero Section - Projects at the Core */}
      <div className={classNames.heroCard}>
        <div className={classNames.heroTitle}>Projects Overview</div>
        <div className={classNames.heroSubtitle}>
          Projects are at the heart of everything - Materials, Procurement, and Finance flow from projects
        </div>
        <div className={classNames.heroStats}>
          <div className={classNames.heroStat}>
            <span className={classNames.heroStatValue}>{metrics?.totalProjects || 0}</span>
            <span className={classNames.heroStatLabel}>Total Projects</span>
          </div>
          <div className={classNames.heroStat}>
            <span className={classNames.heroStatValue}>{metrics?.activeProjects || 0}</span>
            <span className={classNames.heroStatLabel}>Active</span>
          </div>
          <div className={classNames.heroStat}>
            <span className={classNames.heroStatValue}>{metrics?.completedProjects || 0}</span>
            <span className={classNames.heroStatLabel}>Completed</span>
          </div>
        </div>
      </div>

      {/* Key Metrics Grid */}
      <div>
        <Text variant="large" className={classNames.sectionTitle}>
          Quick Overview
        </Text>
        <div className={classNames.metricsGrid}>
          {/* Project-related Metrics */}
          <div className={classNames.metricCard} onClick={() => onNavigate?.('projects')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.blue }}>📁</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.blue }}>
              {metrics?.activeProjects || 0}
            </div>
            <div className={classNames.metricLabel}>Active Projects</div>
          </div>

          {/* Procurement Metrics */}
          <div className={classNames.metricCard} onClick={() => onNavigate?.('procurement')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.orange }}>📦</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.orange }}>
              {metrics?.pendingMaterialRequests || 0}
            </div>
            <div className={classNames.metricLabel}>Pending Material Requests</div>
            {(metrics?.pendingMaterialRequests || 0) > 0 && (
              <div className={classNames.metricChange}>
                <span className={classNames.statusBadge} style={{ backgroundColor: theme.palette.orangeLight, color: theme.palette.orange }}>
                  Needs Attention
                </span>
              </div>
            )}
          </div>

          <div className={classNames.metricCard} onClick={() => onNavigate?.('procurement')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.purple }}>🛒</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.purple }}>
              {metrics?.pendingPurchaseOrders || 0}
            </div>
            <div className={classNames.metricLabel}>Pending Purchase Orders</div>
          </div>

          {/* Finance Metrics */}
          <div className={classNames.metricCard} onClick={() => onNavigate?.('finance')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.green }}>💳</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.green }}>
              {formatCurrency(metrics?.totalPaymentAmount || 0)}
            </div>
            <div className={classNames.metricLabel}>Approved Payments</div>
          </div>

          <div className={classNames.metricCard} onClick={() => onNavigate?.('finance')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.orange }}>⏳</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.orange }}>
              {metrics?.pendingPaymentRequests || 0}
            </div>
            <div className={classNames.metricLabel}>Pending Payments</div>
          </div>

          {/* Materials */}
          <div className={classNames.metricCard} onClick={() => onNavigate?.('material')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.teal }}>📋</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.teal }}>
              {metrics?.totalMaterials || 0}
            </div>
            <div className={classNames.metricLabel}>Materials</div>
          </div>

          {/* Material Management / Inventory */}
          <div className={classNames.metricCard} onClick={() => onNavigate?.('material')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.teal }}>📦</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.teal }}>
              {metrics?.totalInventoryItems || 0}
            </div>
            <div className={classNames.metricLabel}>Inventory Items</div>
          </div>

          <div className={classNames.metricCard} onClick={() => onNavigate?.('material')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.orange }}>⚠️</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.orange }}>
              {metrics?.lowStockItems || 0}
            </div>
            <div className={classNames.metricLabel}>Low Stock Items</div>
            {(metrics?.lowStockItems || 0) > 0 && (
              <div className={classNames.metricChange}>
                <span className={classNames.statusBadge} style={{ backgroundColor: theme.palette.orangeLight, color: theme.palette.orange }}>
                  Needs Attention
                </span>
              </div>
            )}
          </div>

          <div className={classNames.metricCard} onClick={() => onNavigate?.('material')}>
            <div className={classNames.metricIcon} style={{ color: theme.palette.blue }}>🔄</div>
            <div className={classNames.metricValue} style={{ color: theme.palette.blue }}>
              {metrics?.totalMovements || 0}
            </div>
            <div className={classNames.metricLabel}>Inventory Movements</div>
          </div>
        </div>
      </div>

      {/* Charts Section */}
      <div className={classNames.chartsGrid}>
        {/* Project Status Chart */}
        <div className={classNames.chartCard}>
          <Text variant="large" className={classNames.chartTitle}>
            Project Status Distribution
          </Text>
          <DashboardDonutChart
            data={[
              { label: 'Active', value: metrics?.activeProjects || 0, color: theme.palette.blue },
              { label: 'Completed', value: metrics?.completedProjects || 0, color: theme.palette.green },
              { label: 'Other', value: Math.max(0, (metrics?.totalProjects || 0) - (metrics?.activeProjects || 0) - (metrics?.completedProjects || 0)), color: theme.palette.neutralTertiary },
            ]}
            total={metrics?.totalProjects || 0}
          />
        </div>

        {/* Procurement Pipeline */}
        <div className={classNames.chartCard}>
          <Text variant="large" className={classNames.chartTitle}>
            Procurement Pipeline
          </Text>
          <DashboardBarChart
            data={[
              { label: 'MR', value: metrics?.totalMaterialRequests || 0, color: theme.palette.orange },
              { label: 'PR', value: metrics?.totalPurchaseRequisitions || 0, color: theme.palette.yellow },
              { label: 'PO', value: metrics?.totalPurchaseOrders || 0, color: theme.palette.purple },
              { label: 'GRN', value: metrics?.totalGoodsReceived || 0, color: theme.palette.green },
            ]}
            height={220}
          />
        </div>

        {/* Finance Overview */}
        <div className={classNames.chartCard}>
          <Text variant="large" className={classNames.chartTitle}>
            Finance Overview
          </Text>
          <DashboardBarChart
            data={[
              { label: 'Approved', value: metrics?.totalApprovedPayments || 0, color: theme.palette.green },
              { label: 'Pending', value: metrics?.pendingPaymentRequests || 0, color: theme.palette.orange },
              { label: 'Expenses', value: metrics?.totalExpenses || 0, color: theme.palette.blue },
            ]}
            height={220}
          />
        </div>

        {/* Monthly Trend */}
        <div className={classNames.chartCard}>
          <Text variant="large" className={classNames.chartTitle}>
            Activity Trend
          </Text>
          <DashboardLineChart
            data={[
              { label: 'Jan', value: 20 },
              { label: 'Feb', value: 35 },
              { label: 'Mar', value: 45 },
              { label: 'Apr', value: 30 },
              { label: 'May', value: 55 },
              { label: 'Jun', value: 48 },
            ]}
            height={220}
          />
        </div>
      </div>

      {/* Quick Actions */}
      <div>
        <Text variant="large" className={classNames.sectionTitle}>
          Quick Actions
        </Text>
        <div className={classNames.quickActionsGrid}>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('projects')}
          >
            ➕ New Project
          </button>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('procurement')}
          >
            📦 Material Request
          </button>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('procurement')}
          >
            🛒 Purchase Order
          </button>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('finance')}
          >
            💳 Payment Request
          </button>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('material')}
          >
            📋 Materials
          </button>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('material')}
          >
            📦 Material & Inventory
          </button>
          <button
            className={classNames.quickActionButton}
            onClick={() => onNavigate?.('reports')}
          >
            📊 View Reports
          </button>
        </div>
      </div>

      {/* Recent Activity */}
      <div className={classNames.activityCard}>
        <Text variant="large" className={classNames.sectionTitle}>
          Recent Activity
        </Text>
        {recentActivities.length > 0 ? (
          <div>
            {recentActivities.map((activity) => (
              <div
                key={activity.id}
                className={classNames.activityItem}
                onClick={() => {
                  if (activity.type === 'project') onNavigate?.('projects');
                  else if (activity.type === 'procurement') onNavigate?.('procurement');
                  else if (activity.type === 'finance') onNavigate?.('finance');
                  else if (activity.type === 'material') onNavigate?.('material');
                  else if (activity.type === 'inventory') onNavigate?.('material');
                }}
              >
                <div
                  className={classNames.activityIcon}
                  style={{ backgroundColor: `${getActivityColor(activity.type)}20` }}
                >
                  <span style={{ fontSize: '20px' }}>{getActivityIcon(activity.type)}</span>
                </div>
                <div className={classNames.activityContent}>
                  <Text variant="small" block style={{ fontWeight: 600 }}>
                    {activity.title}
                  </Text>
                  <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                    {activity.description}
                  </Text>
                </div>
                <Text variant="small" style={{ color: theme.palette.neutralTertiary }}>
                  {formatTimestamp(activity.timestamp)}
                </Text>
              </div>
            ))}
          </div>
        ) : (
          <div className={classNames.emptyState}>
            <div className={classNames.emptyIcon}>📭</div>
            <div className={classNames.emptyTitle}>No Recent Activity</div>
            <div className={classNames.emptySubtitle}>Start creating projects and requests to see activity here</div>
          </div>
        )}
      </div>

      {/* Summary Stats Row */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))',
        gap: '16px',
      }}>
        <div className={classNames.summaryCard}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '32px', fontWeight: 700, color: theme.palette.blue }}>
              {metrics?.totalVendors || 0}
            </div>
            <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
              Vendors
            </div>
          </div>
        </div>

        <div className={classNames.summaryCard}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '32px', fontWeight: 700, color: theme.palette.purple }}>
              {metrics?.totalPurchaseOrders || 0}
            </div>
            <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
              Total POs
            </div>
          </div>
        </div>

        <div className={classNames.summaryCard}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '32px', fontWeight: 700, color: theme.palette.green }}>
              {metrics?.totalApprovedPayments || 0}
            </div>
            <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
              Approved Payments
            </div>
          </div>
        </div>

        <div className={classNames.summaryCard}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '32px', fontWeight: 700, color: theme.palette.orange }}>
              {metrics?.totalGoodsReceived || 0}
            </div>
            <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
              Goods Received
            </div>
          </div>
        </div>
      </div>
        </>
      )}
    </div>
  );
};

export default DashboardModule;
