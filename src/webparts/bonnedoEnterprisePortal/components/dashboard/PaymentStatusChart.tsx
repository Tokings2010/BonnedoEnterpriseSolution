import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IPaymentStatusChartProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IStatusData {
    status: string;
    count: number;
    amount: number;
}

const PaymentStatusChart: React.FC<IPaymentStatusChartProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [statusData, setStatusData] = React.useState<IStatusData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const response = await spHttpClient.get(`${webUrl}/_api/web/lists/getByTitle('FIN_Payment_Request_Register')/items?$top=500`, SPHttpClient.configurations.v1);

                if (!response.ok) {
                    setStatusData([]);
                    setIsLoading(false);
                    return;
                }

                const data = await response.json();
                const items = data.value || [];

                // Group by status
                const statusMap = new Map<string, { count: number; amount: number }>();

                items.forEach((item: any) => {
                    const status = item.Status || item.Payment_Status || 'Pending';
                    const amount = parseFloat(item.Amount || item.Payment_Amount || '0') || 0;

                    if (statusMap.has(status)) {
                        const existing = statusMap.get(status)!;
                        statusMap.set(status, { count: existing.count + 1, amount: existing.amount + amount });
                    } else {
                        statusMap.set(status, { count: 1, amount });
                    }
                });

                const statuses: IStatusData[] = Array.from(statusMap.entries()).map(([status, data]) => ({
                    status,
                    count: data.count,
                    amount: data.amount,
                }));

                setStatusData(statuses);
            } catch (error) {
                console.error('Error fetching payment status data:', error);
                setStatusData([]);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData();
    }, [refreshKey, pageContext]);

    const total = statusData.reduce((sum, d) => sum + d.count, 0);

    const colors: Record<string, string> = {
        'Pending': theme.palette.orange,
        'Approved': theme.palette.green,
        'Rejected': theme.palette.red,
        'Paid': theme.palette.blue,
        'Partially Paid': theme.palette.purple,
    };

    const defaultColors = [
        theme.palette.blue,
        theme.palette.green,
        theme.palette.orange,
        theme.palette.red,
        theme.palette.purple,
    ];

    // Calculate donut chart segments
    let currentAngle = 0;
    const radius = 70;
    const centerX = 100;
    const centerY = 100;

    const getCoordinatesForPercent = (percent: number): [number, number] => {
        const x = centerX + radius * Math.cos(2 * Math.PI * percent);
        const y = centerY + radius * Math.sin(2 * Math.PI * percent);
        return [x, y];
    };

    if (isLoading) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                Loading payment status data...
            </div>
        );
    }

    if (statusData.length === 0 || total === 0) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                No payment data available
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '20px' }}>
            {/* Donut Chart */}
            <div style={{ position: 'relative', width: '200px', height: '200px' }}>
                <svg viewBox="0 0 200 200" style={{ width: '100%', height: '100%' }}>
                    {statusData.map((item, index) => {
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
                                style={{ cursor: 'pointer', transition: 'opacity 0.2s' }}
                            >
                                <title>{item.status}: {item.count} ({((percent * 100).toFixed(1))}%)</title>
                            </path>
                        );
                    })}
                    {/* Center circle for donut effect */}
                    <circle cx={centerX} cy={centerY} r={40} fill={theme.palette.white} />
                </svg>
                {/* Center text */}
                <div style={{
                    position: 'absolute',
                    top: '50%',
                    left: '50%',
                    transform: 'translate(-50%, -50%)',
                    textAlign: 'center',
                }}>
                    <div style={{ fontSize: '24px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
                        {total}
                    </div>
                    <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                        Total
                    </div>
                </div>
            </div>

            {/* Legend */}
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', justifyContent: 'center' }}>
                {statusData.map((item, index) => {
                    const color = colors[item.status] || defaultColors[index % defaultColors.length];
                    const percent = ((item.count / total) * 100).toFixed(1);
                    return (
                        <div key={item.status} style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <div style={{ width: '12px', height: '12px', borderRadius: '2px', backgroundColor: color }} />
                            <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                                {item.status} ({percent}%)
                            </span>
                        </div>
                    );
                })}
            </div>

            {/* Summary Cards */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '12px', width: '100%' }}>
                {statusData.slice(0, 4).map((item, index) => {
                    const color = colors[item.status] || defaultColors[index % defaultColors.length];
                    return (
                        <div key={item.status} style={{
                            padding: '12px',
                            backgroundColor: `${color}15`,
                            borderRadius: '8px',
                            borderLeft: `3px solid ${color}`,
                        }}>
                            <div style={{ fontSize: '18px', fontWeight: 600, color: theme.palette.neutralPrimary }}>
                                {item.count}
                            </div>
                            <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                                {item.status}
                            </div>
                            <div style={{ fontSize: '11px', color: theme.palette.neutralTertiary }}>
                                ${item.amount.toLocaleString()}
                            </div>
                        </div>
                    );
                })}
            </div>
        </div>
    );
};

export default PaymentStatusChart;
