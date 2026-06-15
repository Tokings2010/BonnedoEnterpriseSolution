import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    Icon,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';

// Budget Tracking Component
interface IFinanceBudgetTrackingProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const FinanceBudgetTracking: React.FC<IFinanceBudgetTrackingProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [projects, setProjects] = React.useState<IListItem[]>([]);
    const [purchaseOrders, setPurchaseOrders] = React.useState<IListItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [refreshKey, setRefreshKey] = React.useState(0);

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    const fetchData = React.useCallback(async () => {
        setLoading(true);
        try {
            const projectsData = await sharePointService.getListData('ENT_Project_Master', undefined, 100);
            const poData = await sharePointService.getListData('PRC_Purchase_Order_Register', undefined, 500);

            setProjects(projectsData);
            setPurchaseOrders(poData);
        } catch (error) {
            console.error('Error fetching budget data:', error);
        }
        setLoading(false);
    }, [sharePointService]);

    React.useEffect(() => {
        fetchData();
    }, [fetchData, refreshKey]);

    const handleRefresh = (): void => {
        setRefreshKey(prev => prev + 1);
        if (onRefresh) onRefresh();
    };

    // Calculate total procurement cost per project
    const getProjectProcurementCost = (projectCode: string): number => {
        return purchaseOrders
            .filter(po => po.Project_Code === projectCode)
            .reduce((sum, po) => sum + (po.Amount || po.TotalAmount || 0), 0);
    };

    // Calculate remaining budget
    const getRemainingBudget = (project: IListItem): number => {
        const contractValue = project.Contract_Value || project.Budget || 0;
        const procurementCost = getProjectProcurementCost(project.Project_Code || project.Title);
        return contractValue - procurementCost;
    };

    // Calculate utilization percentage
    const getUtilizationPercentage = (project: IListItem): number => {
        const contractValue = project.Contract_Value || project.Budget || 0;
        if (contractValue === 0) return 0;
        const procurementCost = getProjectProcurementCost(project.Project_Code || project.Title);
        return (procurementCost / contractValue) * 100;
    };

    // Get budget status color
    const getBudgetStatusColor = (remaining: number, total: number): string => {
        const percentage = (remaining / total) * 100;
        if (percentage > 50) return '#107c10'; // Green - good
        if (percentage > 20) return '#ca5010'; // Orange - warning
        return '#a80000'; // Red - critical
    };

    const classNames = mergeStyleSets({
        container: { padding: '20px', height: '100%' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' },
        headerTitle: { flex: 1, minWidth: '200px' },
        headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
        summaryCards: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
            gap: '16px',
            marginBottom: '24px'
        },
        summaryCard: {
            padding: '20px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            boxShadow: theme.effects.elevation4,
            border: `1px solid ${theme.palette.neutralLight}`,
        },
        summaryTitle: {
            fontSize: '14px',
            color: theme.palette.neutralSecondary,
            marginBottom: '8px',
        },
        summaryValue: {
            fontSize: '24px',
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
        },
        tableContainer: {
            overflowY: 'auto',
            maxHeight: 'calc(100vh - 400px)',
        },
        table: {
            width: '100%',
            borderCollapse: 'collapse',
        },
        th: {
            textAlign: 'left',
            padding: '12px',
            backgroundColor: theme.palette.neutralLighterAlt,
            borderBottom: `2px solid ${theme.palette.neutralLight}`,
            fontWeight: 600,
            position: 'sticky',
            top: 0,
        },
        td: {
            padding: '12px',
            borderBottom: `1px solid ${theme.palette.neutralLight}`,
        },
        progressBar: {
            height: '8px',
            backgroundColor: theme.palette.neutralLighter,
            borderRadius: '4px',
            overflow: 'hidden',
        },
        progressFill: {
            height: '100%',
            borderRadius: '4px',
            transition: 'width 0.3s ease',
        },
        loadingContainer: {
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            height: '200px',
            color: theme.palette.neutralSecondary,
        },
    });

    // Calculate totals
    const totalContractValue = projects.reduce((sum, p) => sum + (p.Contract_Value || p.Budget || 0), 0);
    const totalProcurementCost = projects.reduce((sum, p) => {
        const projectCode = p.Project_Code || p.Title;
        return sum + getProjectProcurementCost(projectCode);
    }, 0);
    const totalRemainingBudget = totalContractValue - totalProcurementCost;

    return (
        <div className={classNames.container}>
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Budget Tracking
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Monitor and analyze budget allocations
                    </Text>
                </div>
                <div className={classNames.headerActions}>
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={handleRefresh}
                        title="Refresh data"
                        ariaLabel="Refresh data"
                    />
                </div>
            </div>

            {/* Summary Cards */}
            <div className={classNames.summaryCards}>
                <div className={classNames.summaryCard}>
                    <div className={classNames.summaryTitle}>Total Contract Value</div>
                    <div className={classNames.summaryValue}>
                        {totalContractValue.toLocaleString('en-NG', { style: 'currency', currency: 'NGN' })}
                    </div>
                </div>
                <div className={classNames.summaryCard}>
                    <div className={classNames.summaryTitle}>Total Procurement Cost</div>
                    <div className={classNames.summaryValue}>
                        {totalProcurementCost.toLocaleString('en-NG', { style: 'currency', currency: 'NGN' })}
                    </div>
                </div>
                <div className={classNames.summaryCard}>
                    <div className={classNames.summaryTitle}>Total Remaining Budget</div>
                    <div className={classNames.summaryValue} style={{ color: getBudgetStatusColor(totalRemainingBudget, totalContractValue) }}>
                        {totalRemainingBudget.toLocaleString('en-NG', { style: 'currency', currency: 'NGN' })}
                    </div>
                </div>
                <div className={classNames.summaryCard}>
                    <div className={classNames.summaryTitle}>Budget Utilization</div>
                    <div className={classNames.summaryValue}>
                        {totalContractValue > 0 ? ((totalProcurementCost / totalContractValue) * 100).toFixed(1) : 0}%
                    </div>
                </div>
            </div>

            {loading ? (
                <div className={classNames.loadingContainer}>
                    <Icon iconName="Sync" style={{ fontSize: '24px', animation: 'spin 1s linear infinite' }} />
                    <span style={{ marginLeft: '8px' }}>Loading budget data...</span>
                </div>
            ) : (
                <div className={classNames.tableContainer}>
                    <table className={classNames.table}>
                        <thead>
                            <tr>
                                <th className={classNames.th}>Project Code</th>
                                <th className={classNames.th}>Project Name</th>
                                <th className={classNames.th}>Contract Value</th>
                                <th className={classNames.th}>Procurement Cost</th>
                                <th className={classNames.th}>Remaining Budget</th>
                                <th className={classNames.th}>Utilization</th>
                            </tr>
                        </thead>
                        <tbody>
                            {projects.map((project) => {
                                const projectCode = project.Project_Code || project.Title;
                                const contractValue = project.Contract_Value || project.Budget || 0;
                                const procurementCost = getProjectProcurementCost(projectCode);
                                const remaining = getRemainingBudget(project);
                                const utilization = getUtilizationPercentage(project);

                                return (
                                    <tr key={project.ID}>
                                        <td className={classNames.td}>{projectCode}</td>
                                        <td className={classNames.td}>{project.Project_Name || project.Title}</td>
                                        <td className={classNames.td}>
                                            {contractValue.toLocaleString('en-NG', { style: 'currency', currency: 'NGN' })}
                                        </td>
                                        <td className={classNames.td}>
                                            {procurementCost.toLocaleString('en-NG', { style: 'currency', currency: 'NGN' })}
                                        </td>
                                        <td className={classNames.td} style={{ color: getBudgetStatusColor(remaining, contractValue), fontWeight: 600 }}>
                                            {remaining.toLocaleString('en-NG', { style: 'currency', currency: 'NGN' })}
                                        </td>
                                        <td className={classNames.td}>
                                            <div className={classNames.progressBar}>
                                                <div
                                                    className={classNames.progressFill}
                                                    style={{
                                                        width: `${Math.min(utilization, 100)}%`,
                                                        backgroundColor: getBudgetStatusColor(remaining, contractValue)
                                                    }}
                                                />
                                            </div>
                                            <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                                                {utilization.toFixed(1)}%
                                            </span>
                                        </td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
};

export default FinanceBudgetTracking;
