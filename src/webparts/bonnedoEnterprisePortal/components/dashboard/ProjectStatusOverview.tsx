import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IProjectStatusOverviewProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IProjectStatusData {
    status: string;
    count: number;
    budget: number;
    spent: number;
}

const ProjectStatusOverview: React.FC<IProjectStatusOverviewProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [projectData, setProjectData] = React.useState<IProjectStatusData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);
    const [totalBudget, setTotalBudget] = React.useState(0);
    const [totalSpent, setTotalSpent] = React.useState(0);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const response = await spHttpClient.get(`${webUrl}/_api/web/lists/getByTitle('ENT_Project_Master')/items?$top=500`, SPHttpClient.configurations.v1);

                if (!response.ok) {
                    setProjectData([]);
                    setIsLoading(false);
                    return;
                }

                const data = await response.json();
                const items = data.value || [];

                // Group by status
                const statusMap = new Map<string, { count: number; budget: number; spent: number }>();

                items.forEach((item: any) => {
                    const status = item.Status || item.Project_Status || 'Active';
                    const budget = parseFloat(item.Contract_Value || item.Budget || item.Project_Budget || '0') || 0;
                    const spent = parseFloat(item.Total_Spent || item.Spent_Amount || '0') || 0;

                    if (statusMap.has(status)) {
                        const existing = statusMap.get(status)!;
                        statusMap.set(status, {
                            count: existing.count + 1,
                            budget: existing.budget + budget,
                            spent: existing.spent + spent,
                        });
                    } else {
                        statusMap.set(status, { count: 1, budget, spent });
                    }
                });

                const statuses: IProjectStatusData[] = Array.from(statusMap.entries()).map(([status, data]) => ({
                    status,
                    count: data.count,
                    budget: data.budget,
                    spent: data.spent,
                }));

                setProjectData(statuses);
                setTotalBudget(statuses.reduce((sum, s) => sum + s.budget, 0));
                setTotalSpent(statuses.reduce((sum, s) => sum + s.spent, 0));
            } catch (error) {
                console.error('Error fetching project status data:', error);
                setProjectData([]);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData();
    }, [refreshKey, pageContext]);

    const total = projectData.reduce((sum, d) => sum + d.count, 0);

    const colors: Record<string, string> = {
        'Active': theme.palette.green,
        'Completed': theme.palette.blue,
        'On Hold': theme.palette.orange,
        'Cancelled': theme.palette.red,
        'Planning': theme.palette.purple,
    };

    const defaultColors = [
        theme.palette.green,
        theme.palette.blue,
        theme.palette.orange,
        theme.palette.red,
        theme.palette.purple,
    ];

    // Calculate pie chart segments
    let currentAngle = 0;
    const radius = 60;
    const centerX = 80;
    const centerY = 80;

    const getCoordinatesForPercent = (percent: number): [number, number] => {
        const x = centerX + radius * Math.cos(2 * Math.PI * percent);
        const y = centerY + radius * Math.sin(2 * Math.PI * percent);
        return [x, y];
    };

    if (isLoading) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                Loading project status data...
            </div>
        );
    }

    if (projectData.length === 0 || total === 0) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                No project data available
            </div>
        );
    }

    // Calculate budget utilization
    const budgetUtilization = totalBudget > 0 ? (totalSpent / totalBudget) * 100 : 0;

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            {/* Budget Overview */}
            <div style={{
                padding: '16px',
                backgroundColor: theme.palette.neutralLighterAlt,
                borderRadius: '8px',
            }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px' }}>
                    <span style={{ fontSize: '14px', color: theme.palette.neutralSecondary }}>
                        Total Budget Utilization
                    </span>
                    <span style={{ fontSize: '14px', fontWeight: 600, color: theme.palette.neutralPrimary }}>
                        {budgetUtilization.toFixed(1)}%
                    </span>
                </div>
                <div style={{
                    height: '8px',
                    backgroundColor: theme.palette.neutralLight,
                    borderRadius: '4px',
                    overflow: 'hidden',
                }}>
                    <div style={{
                        height: '100%',
                        width: `${Math.min(budgetUtilization, 100)}%`,
                        backgroundColor: budgetUtilization > 90 ? theme.palette.red : budgetUtilization > 70 ? theme.palette.orange : theme.palette.green,
                        borderRadius: '4px',
                    }} />
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '8px', fontSize: '12px' }}>
                    <span style={{ color: theme.palette.neutralSecondary }}>
                        Spent: ${totalSpent.toLocaleString()}
                    </span>
                    <span style={{ color: theme.palette.neutralSecondary }}>
                        Budget: ${totalBudget.toLocaleString()}
                    </span>
                </div>
            </div>

            {/* Project Status Distribution */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '20px', flexWrap: 'wrap' }}>
                {/* Pie Chart */}
                <div style={{ position: 'relative', width: '160px', height: '160px', flexShrink: 0 }}>
                    <svg viewBox="0 0 160 160" style={{ width: '100%', height: '100%' }}>
                        {projectData.map((item, index) => {
                            const percent = item.count / total;
                            const startAngle = currentAngle;
                            currentAngle += percent;

                            const [startX, startY] = getCoordinatesForPercent(startAngle);
                            const [endX, endY] = getCoordinatesForPercent(currentAngle);

                            const largeArcFlag = percent > 0.5 ? 1 : 0;

                            const pathData = [
                                `M ${centerX} ${centerY}`,
                                `L ${startX} ${startY}`,
                                `A ${radius} ${radius} 0 ${largeArcFlag} 1 ${endX} ${endY}`,
                                'Z',
                            ].join(' ');

                            const color = colors[item.status] || defaultColors[index % defaultColors.length];

                            return (
                                <path
                                    key={item.status}
                                    d={pathData}
                                    fill={color}
                                    style={{ cursor: 'pointer' }}
                                >
                                    <title>{item.status}: {item.count} projects</title>
                                </path>
                            );
                        })}
                        <circle cx={centerX} cy={centerY} r={25} fill={theme.palette.white} />
                    </svg>
                    <div style={{
                        position: 'absolute',
                        top: '50%',
                        left: '50%',
                        transform: 'translate(-50%, -50%)',
                        textAlign: 'center',
                    }}>
                        <div style={{ fontSize: '20px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
                            {total}
                        </div>
                        <div style={{ fontSize: '10px', color: theme.palette.neutralSecondary }}>
                            Projects
                        </div>
                    </div>
                </div>

                {/* Legend & Details */}
                <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '8px' }}>
                    {projectData.map((item, index) => {
                        const color = colors[item.status] || defaultColors[index % defaultColors.length];
                        const percent = ((item.count / total) * 100).toFixed(1);
                        return (
                            <div key={item.status} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <div style={{
                                    width: '10px',
                                    height: '10px',
                                    borderRadius: '2px',
                                    backgroundColor: color,
                                    flexShrink: 0,
                                }} />
                                <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{
                                        fontSize: '13px',
                                        fontWeight: 500,
                                        color: theme.palette.neutralPrimary,
                                        whiteSpace: 'nowrap',
                                        overflow: 'hidden',
                                        textOverflow: 'ellipsis',
                                    }}>
                                        {item.status}
                                    </div>
                                </div>
                                <div style={{ fontSize: '13px', color: theme.palette.neutralSecondary }}>
                                    {item.count} ({percent}%)
                                </div>
                            </div>
                        );
                    })}
                </div>
            </div>
        </div>
    );
};

export default ProjectStatusOverview;
