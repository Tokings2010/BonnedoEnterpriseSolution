import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    Icon,
    DetailsList,
    DetailsListLayoutMode,
    SelectionMode,
    IColumn,
    Spinner,
    SpinnerSize,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';

interface IFinanceBudgetTrackingProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

interface IBudgetRecord {
    ID: number;
    Project_Code: string;
    Project_Name: string;
    Contract_Value: number;
    Procurement_Cost: number;
    Remaining_Budget: number;
    Utilization: number;
    Budget_Status: string;
}

const FinanceBudgetTracking: React.FC<IFinanceBudgetTrackingProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [budgetRecords, setBudgetRecords] = React.useState<IBudgetRecord[]>([]);
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

            // Compute procurement cost per project
            const getProjectProcurementCost = (projectCode: string): number => {
                return poData
                    .filter(po => {
                        const poProject = po.Project_Code || po.Project_x0020_Code || '';
                        return String(poProject).trim() === String(projectCode).trim();
                    })
                    .reduce((sum, po) => sum + (po.TotalAmount || po.Amount || 0), 0);
            };

            const getBudgetStatusColor = (remaining: number, total: number): string => {
                if (total === 0) return '#107c10';
                const percentage = (remaining / total) * 100;
                if (percentage > 50) return '#107c10';
                if (percentage > 20) return '#ca5010';
                return '#a80000';
            };

            const records: IBudgetRecord[] = projectsData.map((project) => {
                const projectCode = project.Project_Code || project.Project_x0020_Code || project.Title || '';
                const contractValue = project.Contract_Value || project.Budget || 0;
                const procurementCost = getProjectProcurementCost(projectCode);
                const remaining = contractValue - procurementCost;
                const utilization = contractValue > 0 ? (procurementCost / contractValue) * 100 : 0;
                return {
                    ID: project.ID,
                    Project_Code: projectCode,
                    Project_Name: project.Project_Name || project.Title || '',
                    Contract_Value: contractValue,
                    Procurement_Cost: procurementCost,
                    Remaining_Budget: remaining,
                    Utilization: utilization,
                    Budget_Status: remaining <= 0 ? 'Exhausted' : remaining < contractValue * 0.2 ? 'Critical' : remaining < contractValue * 0.5 ? 'Warning' : 'Healthy',
                };
            });

            setBudgetRecords(records);
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

    const fmtCur = (n: number): string => {
        if (n >= 1000000) return `\u20A6${(n / 1000000).toFixed(1)}M`;
        if (n >= 1000) return `\u20A6${(n / 1000).toFixed(1)}K`;
        return `\u20A6${n.toFixed(0)}`;
    };

    const getBudgetStatusColor = (remaining: number, total: number): string => {
        if (total === 0) return '#107c10';
        const percentage = (remaining / total) * 100;
        if (percentage > 50) return '#107c10';
        if (percentage > 20) return '#ca5010';
        return '#a80000';
    };

    // Compute totals
    const totalContractValue = budgetRecords.reduce((sum, r) => sum + r.Contract_Value, 0);
    const totalProcurementCost = budgetRecords.reduce((sum, r) => sum + r.Procurement_Cost, 0);
    const totalRemainingBudget = totalContractValue - totalProcurementCost;

    const budgetColumns: IColumn[] = [
        {
            key: 'Project_Code',
            name: 'Project Code',
            fieldName: 'Project_Code',
            minWidth: 130,
            isResizable: true,
        },
        {
            key: 'Project_Name',
            name: 'Project Name',
            fieldName: 'Project_Name',
            minWidth: 200,
            isResizable: true,
        },
        {
            key: 'Contract_Value',
            name: 'Contract Value',
            fieldName: 'Contract_Value',
            minWidth: 150,
            isResizable: true,
            onRender: (item?: IBudgetRecord) => item ? (
                <span style={{ fontVariantNumeric: 'tabular-nums', fontWeight: 500 }}>
                    {fmtCur(item.Contract_Value)}
                </span>
            ) : null,
        },
        {
            key: 'Procurement_Cost',
            name: 'Procurement Cost',
            fieldName: 'Procurement_Cost',
            minWidth: 150,
            isResizable: true,
            onRender: (item?: IBudgetRecord) => item ? (
                <span style={{ fontVariantNumeric: 'tabular-nums', fontWeight: 500, color: '#D13438' }}>
                    {fmtCur(item.Procurement_Cost)}
                </span>
            ) : null,
        },
        {
            key: 'Remaining_Budget',
            name: 'Remaining Budget',
            fieldName: 'Remaining_Budget',
            minWidth: 150,
            isResizable: true,
            onRender: (item?: IBudgetRecord) => item ? (
                <span style={{
                    fontVariantNumeric: 'tabular-nums',
                    fontWeight: 600,
                    color: getBudgetStatusColor(item.Remaining_Budget, item.Contract_Value),
                }}>
                    {fmtCur(item.Remaining_Budget)}
                </span>
            ) : null,
        },
        {
            key: 'Utilization',
            name: 'Utilization',
            fieldName: 'Utilization',
            minWidth: 160,
            isResizable: true,
            onRender: (item?: IBudgetRecord) => {
                if (!item) return null;
                const pct = Math.min(item.Utilization, 100);
                const color = getBudgetStatusColor(item.Remaining_Budget, item.Contract_Value);
                return (
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, width: '100%' }}>
                        <span style={{
                            flex: 1, height: 8, backgroundColor: '#EDEDED', borderRadius: 4,
                            overflow: 'hidden', display: 'inline-block', maxWidth: 120,
                        }}>
                            <span style={{
                                display: 'block', height: 8, borderRadius: 4,
                                width: `${pct}%`, backgroundColor: color,
                                transition: 'width 0.3s ease',
                            }} />
                        </span>
                        <span style={{ fontSize: 12, color: '#616161', fontWeight: 500 }}>
                            {pct.toFixed(1)}%
                        </span>
                    </span>
                );
            },
        },
    ];

    const classNames = mergeStyleSets({
        container: { padding: '20px', height: '100%', display: 'flex', flexDirection: 'column' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' },
        headerTitle: { flex: 1, minWidth: '200px' },
        headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
        summaryCards: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
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
            fontSize: '13px',
            color: theme.palette.neutralSecondary,
            marginBottom: '8px',
        },
        summaryValue: {
            fontSize: '24px',
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
        },
        gridContainer: {
            flex: 1,
            overflow: 'auto',
            backgroundColor: theme.palette.white,
            border: `1px solid ${theme.palette.neutralLight}`,
            borderRadius: '4px',
        },
        loadingContainer: {
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            height: '200px',
            color: theme.palette.neutralSecondary,
        },
    });

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
                        {fmtCur(totalContractValue)}
                    </div>
                </div>
                <div className={classNames.summaryCard}>
                    <div className={classNames.summaryTitle}>Total Procurement Cost</div>
                    <div className={classNames.summaryValue} style={{ color: '#D13438' }}>
                        {fmtCur(totalProcurementCost)}
                    </div>
                </div>
                <div className={classNames.summaryCard}>
                    <div className={classNames.summaryTitle}>Total Remaining Budget</div>
                    <div className={classNames.summaryValue} style={{ color: getBudgetStatusColor(totalRemainingBudget, totalContractValue) }}>
                        {fmtCur(totalRemainingBudget)}
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
                    <Icon iconName="Sync" style={{ fontSize: '24px' }} />
                    <span style={{ marginLeft: '8px' }}>Loading budget data...</span>
                </div>
            ) : (
                <div className={classNames.gridContainer}>
                    <DetailsList
                        items={budgetRecords}
                        columns={budgetColumns}
                        layoutMode={DetailsListLayoutMode.justified}
                        selectionMode={SelectionMode.none}
                        isHeaderVisible={true}
                        compact={true}
                    />
                </div>
            )}
        </div>
    );
};

export default FinanceBudgetTracking;
