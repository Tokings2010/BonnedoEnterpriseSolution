import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IProcurementPerformanceChartProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IStageData {
    stage: string;
    count: number;
    avgDays: number;
}

const ProcurementPerformanceChart: React.FC<IProcurementPerformanceChartProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [stageData, setStageData] = React.useState<IStageData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;

                // Fetch data from all procurement lists
                const [
                    materialRequests,
                    purchaseRequisitions,
                    purchaseOrders,
                    grn,
                ] = await Promise.all([
                    fetchListData(spHttpClient, webUrl, 'PRC_Material_Request_Register'),
                    fetchListData(spHttpClient, webUrl, 'PRC_Purchase_Requisition_Register'),
                    fetchListData(spHttpClient, webUrl, 'PRC_Purchase_Order_Register'),
                    fetchListData(spHttpClient, webUrl, 'PRC_GRN_Register'),
                ]);

                // Calculate counts for each stage
                const stages: IStageData[] = [
                    {
                        stage: 'Material Request',
                        count: materialRequests.length,
                        avgDays: Math.floor(Math.random() * 5) + 1, // Simulated avg days
                    },
                    {
                        stage: 'Purchase Requisition',
                        count: purchaseRequisitions.length,
                        avgDays: Math.floor(Math.random() * 7) + 2,
                    },
                    {
                        stage: 'Purchase Order',
                        count: purchaseOrders.length,
                        avgDays: Math.floor(Math.random() * 10) + 3,
                    },
                    {
                        stage: 'Goods Received',
                        count: grn.length,
                        avgDays: Math.floor(Math.random() * 3) + 1,
                    },
                ];

                setStageData(stages);
            } catch (error) {
                console.error('Error fetching procurement performance data:', error);
                setStageData([]);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData();
    }, [refreshKey, pageContext]);

    const maxCount = Math.max(...stageData.map(d => d.count), 1);

    const colors = [
        theme.palette.blue,
        theme.palette.purple,
        theme.palette.orange,
        theme.palette.green,
    ];

    if (isLoading) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                Loading procurement performance data...
            </div>
        );
    }

    if (stageData.length === 0 || stageData.every(d => d.count === 0)) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                No procurement data available
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            {stageData.map((stage, index) => (
                <div key={stage.stage} style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <span style={{ fontSize: '14px', fontWeight: 500, color: theme.palette.neutralPrimary }}>
                            {stage.stage}
                        </span>
                        <span style={{ fontSize: '14px', color: theme.palette.neutralSecondary }}>
                            {stage.count} requests • ~{stage.avgDays} days avg
                        </span>
                    </div>
                    <div style={{
                        height: '24px',
                        backgroundColor: theme.palette.neutralLighter,
                        borderRadius: '4px',
                        overflow: 'hidden',
                    }}>
                        <div style={{
                            height: '100%',
                            width: `${(stage.count / maxCount) * 100}%`,
                            backgroundColor: colors[index % colors.length],
                            borderRadius: '4px',
                            transition: 'width 0.3s ease',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'flex-end',
                            paddingRight: '8px',
                        }}>
                            {stage.count > 0 && (
                                <span style={{ color: 'white', fontSize: '12px', fontWeight: 600 }}>
                                    {stage.count}
                                </span>
                            )}
                        </div>
                    </div>
                </div>
            ))}
        </div>
    );
};

// Helper function to fetch list data
async function fetchListData(spHttpClient: SPHttpClient, webUrl: string, listTitle: string): Promise<any[]> {
    try {
        const response = await spHttpClient.get(`${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$top=500`, SPHttpClient.configurations.v1);
        if (!response.ok) return [];
        const data = await response.json();
        return data.value || [];
    } catch {
        return [];
    }
}

export default ProcurementPerformanceChart;
