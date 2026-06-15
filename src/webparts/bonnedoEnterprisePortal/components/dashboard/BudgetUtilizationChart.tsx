import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IBudgetUtilizationChartProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IProjectBudget {
    projectName: string;
    budget: number;
    spent: number;
    remaining: number;
    percentage: number;
}

const BudgetUtilizationChart: React.FC<IBudgetUtilizationChartProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [projectData, setProjectData] = React.useState<IProjectBudget[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;

                // Fetch projects and purchase orders
                const [projects, purchaseOrders] = await Promise.all([
                    fetchListData(webUrl, 'ENT_Project_Master'),
                    fetchListData(webUrl, 'PRC_Purchase_Order_Register'),
                ]);

                // Calculate spend per project from purchase orders
                const projectSpendMap = new Map<string, number>();

                purchaseOrders.forEach((po: any) => {
                    const projectName = po.Project || po.Project_Name || po.Project_Title || 'Unknown';
                    const amount = parseFloat(po.Amount || po.Total_Amount || po.PO_Amount || '0') || 0;

                    if (projectSpendMap.has(projectName)) {
                        projectSpendMap.set(projectName, projectSpendMap.get(projectName)! + amount);
                    } else {
                        projectSpendMap.set(projectName, amount);
                    }
                });

                // Create project budget data
                const budgets: IProjectBudget[] = projects
                    .slice(0, 5) // Top 5 projects
                    .map((project: any) => {
                        const projectName = project.Project_Name || project.Title || project.Project_Code || 'Unknown Project';
                        const budget = parseFloat(project.Contract_Value || project.Budget || project.Project_Budget || '0') || 0;
                        const spent = projectSpendMap.get(projectName) || 0;
                        const remaining = Math.max(0, budget - spent);
                        const percentage = budget > 0 ? Math.min(100, (spent / budget) * 100) : 0;

                        return {
                            projectName,
                            budget,
                            spent,
                            remaining,
                            percentage,
                        };
                    })
                    .filter(p => p.budget > 0) // Only show projects with budget
                    .sort((a, b) => b.percentage - a.percentage); // Sort by utilization

                setProjectData(budgets);
            } catch (error) {
                console.error('Error fetching budget data:', error);
                setProjectData([]);
            }
            setIsLoading(false);
        };

        fetchData();
    }, [spHttpClient, pageContext, refreshKey]);

    const fetchListData = async (webUrl: string, listName: string): Promise<any[]> => {
        try {
            const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=500`;
            const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (!response.ok) return [];
            const data = await response.json();
            return data.value || [];
        } catch {
            return [];
        }
    };

    const formatCurrency = (amount: number): string => {
        if (amount >= 1000000) return `N${(amount / 1000000).toFixed(1)}M`;
        if (amount >= 1000) return `N${(amount / 1000).toFixed(1)}K`;
        return `N${amount.toFixed(0)}`;
    };

    const getProgressColor = (percentage: number): string => {
        if (percentage >= 90) return theme.palette.red;
        if (percentage >= 70) return theme.palette.orange;
        return theme.palette.green;
    };

    if (isLoading) {
        return (
            <div style={{ height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <span style={{ color: theme.palette.neutralSecondary }}>Loading...</span>
            </div>
        );
    }

    if (projectData.length === 0) {
        return (
            <div style={{
                height: '200px',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center'
            }}>
                <span style={{ fontSize: '32px', marginBottom: '8px' }}>📈</span>
                <span style={{ color: theme.palette.neutralSecondary }}>No budget data available</span>
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            {projectData.map((project) => (
                <div key={project.projectName}>
                    <div style={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        marginBottom: '4px'
                    }}>
                        <span style={{
                            fontSize: '12px',
                            fontWeight: 500,
                            color: theme.palette.neutralPrimary
                        }}>
                            {project.projectName.length > 20 ? project.projectName.substring(0, 20) + '...' : project.projectName}
                        </span>
                        <span style={{
                            fontSize: '12px',
                            fontWeight: 600,
                            color: getProgressColor(project.percentage)
                        }}>
                            {project.percentage.toFixed(0)}%
                        </span>
                    </div>

                    {/* Stacked bar */}
                    <div style={{
                        height: '24px',
                        backgroundColor: theme.palette.neutralLight,
                        borderRadius: '4px',
                        overflow: 'hidden',
                        display: 'flex',
                    }}>
                        {/* Spent portion */}
                        <div
                            style={{
                                width: `${project.percentage}%`,
                                backgroundColor: getProgressColor(project.percentage),
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                            }}
                        >
                            {project.percentage > 15 && (
                                <span style={{
                                    fontSize: '10px',
                                    fontWeight: 600,
                                    color: theme.palette.white
                                }}>
                                    {formatCurrency(project.spent)}
                                </span>
                            )}
                        </div>
                        {/* Remaining portion */}
                        {project.remaining > 0 && (
                            <div
                                style={{
                                    flex: 1,
                                    backgroundColor: theme.palette.greenLight,
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                }}
                            >
                                {project.percentage < 85 && (
                                    <span style={{
                                        fontSize: '10px',
                                        fontWeight: 500,
                                        color: theme.palette.green
                                    }}>
                                        {formatCurrency(project.remaining)}
                                    </span>
                                )}
                            </div>
                        )}
                    </div>

                    {/* Budget info */}
                    <div style={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        marginTop: '4px',
                        fontSize: '10px',
                        color: theme.palette.neutralSecondary,
                    }}>
                        <span>Budget: {formatCurrency(project.budget)}</span>
                        <span>Spent: {formatCurrency(project.spent)}</span>
                    </div>
                </div>
            ))}
        </div>
    );
};

export default BudgetUtilizationChart;
