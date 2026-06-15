import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IProcurementTimelineProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IStageData {
    stage: string;
    count: number;
    color: string;
}

const ProcurementTimeline: React.FC<IProcurementTimelineProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [stageData, setStageData] = React.useState<IStageData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);

    const stageColors = [
        theme.palette.orange,
        theme.palette.yellow,
        theme.palette.purple,
        theme.palette.green,
    ];

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;

                // Fetch all procurement stages in parallel
                const [materialRequests, purchaseRequisitions, purchaseOrders, goodsReceived] = await Promise.all([
                    fetchListData(webUrl, 'PRC_Material_Request_Register'),
                    fetchListData(webUrl, 'PRC_Purchase_Requisition_Register'),
                    fetchListData(webUrl, 'PRC_Purchase_Order_Register'),
                    fetchListData(webUrl, 'PRC_GRN_Register'),
                ]);

                const stages: IStageData[] = [
                    { stage: 'Material Requests', count: materialRequests.length, color: stageColors[0] },
                    { stage: 'Purchase Requisitions', count: purchaseRequisitions.length, color: stageColors[1] },
                    { stage: 'Purchase Orders', count: purchaseOrders.length, color: stageColors[2] },
                    { stage: 'Goods Received', count: goodsReceived.length, color: stageColors[3] },
                ];

                setStageData(stages);
            } catch (error) {
                console.error('Error fetching procurement timeline data:', error);
                setStageData([]);
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

    const maxCount = Math.max(...stageData.map(s => s.count), 1);

    if (isLoading) {
        return (
            <div style={{ height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <span style={{ color: theme.palette.neutralSecondary }}>Loading...</span>
            </div>
        );
    }

    if (stageData.length === 0) {
        return (
            <div style={{
                height: '200px',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center'
            }}>
                <span style={{ fontSize: '32px', marginBottom: '8px' }}>📋</span>
                <span style={{ color: theme.palette.neutralSecondary }}>No procurement data available</span>
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            {/* Progress bar showing the pipeline flow */}
            <div style={{
                display: 'flex',
                gap: '4px',
                marginBottom: '8px',
                padding: '0 8px'
            }}>
                {stageData.map((stage, index) => (
                    <div
                        key={stage.stage}
                        style={{
                            flex: 1,
                            height: '6px',
                            backgroundColor: stage.color,
                            borderRadius: index === 0 ? '3px 0 0 3px' : index === stageData.length - 1 ? '0 3px 3px 0' : '0',
                            opacity: stage.count > 0 ? 1 : 0.3,
                        }}
                    />
                ))}
            </div>

            {/* Horizontal bars for each stage */}
            {stageData.map((stage) => (
                <div key={stage.stage} style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <div style={{ width: '120px', fontSize: '12px', fontWeight: 500, color: theme.palette.neutralPrimary, flexShrink: 0 }}>
                        {stage.stage}
                    </div>
                    <div style={{ flex: 1, position: 'relative' }}>
                        <div
                            style={{
                                height: '28px',
                                width: `${(stage.count / maxCount) * 100}%`,
                                backgroundColor: stage.color,
                                borderRadius: '4px',
                                minWidth: '30px',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'flex-end',
                                paddingRight: '8px',
                            }}
                        >
                            <span style={{ fontSize: '12px', fontWeight: 600, color: theme.palette.white }}>
                                {stage.count}
                            </span>
                        </div>
                    </div>
                </div>
            ))}

            {/* Summary */}
            <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                padding: '12px',
                backgroundColor: theme.palette.neutralLighterAlt,
                borderRadius: '8px',
                marginTop: '8px'
            }}>
                <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                    Total Records
                </span>
                <span style={{ fontSize: '14px', fontWeight: 600, color: theme.palette.neutralPrimary }}>
                    {stageData.reduce((sum, s) => sum + s.count, 0)}
                </span>
            </div>
        </div>
    );
};

export default ProcurementTimeline;
