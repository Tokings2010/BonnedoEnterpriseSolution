import * as React from 'react';
import { getTheme } from '@fluentui/react';

interface ILineData {
    label: string;
    value: number;
}

interface IDashboardLineChartProps {
    data: ILineData[];
    height?: number;
}

const DashboardLineChart: React.FC<IDashboardLineChartProps> = ({
    data = [],
    height = 200,
}) => {
    const theme = getTheme();

    if (!data || data.length === 0) {
        return <div style={{ height, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#666' }}>No data</div>;
    }

    // Fixed dimensions for proper aspect ratio
    const chartWidth = 320;
    const chartHeight = 140;
    const padding = 30;
    const maxValue = Math.max(...data.map(d => d.value), 1);
    const minValue = Math.min(...data.map(d => d.value), 0);
    const valueRange = maxValue - minValue || 1;

    // Calculate points for the line
    const points = data.map((item, index) => {
        const x = padding + (index / (data.length - 1)) * (chartWidth - 2 * padding);
        const y = chartHeight - 20 - ((item.value - minValue) / valueRange) * (chartHeight - 40);
        return { x, y, ...item };
    });

    // Create smooth curve path
    const createSmoothPath = (): string => {
        if (points.length < 2) return '';

        let path = `M ${points[0].x} ${points[0].y}`;

        for (let i = 1; i < points.length; i++) {
            const prev = points[i - 1];
            const curr = points[i];
            const cpx = (prev.x + curr.x) / 2;
            path += ` Q ${prev.x + (cpx - prev.x)} ${prev.y}, ${cpx} ${(prev.y + curr.y) / 2}`;
            path += ` Q ${cpx} ${curr.y}, ${curr.x} ${curr.y}`;
        }

        return path;
    };

    // Create area fill path
    const createAreaPath = (): string => {
        if (points.length < 2) return '';

        const areaPath = createSmoothPath();
        const lastPoint = points[points.length - 1];
        const firstPoint = points[0];

        return `${areaPath} L ${lastPoint.x} ${chartHeight - 20} L ${firstPoint.x} ${chartHeight - 20} Z`;
    };

    return (
        <div style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            gap: '12px',
            overflow: 'hidden',
            width: '100%',
        }}>
            {/* Line Chart */}
            <svg width="100%" height={chartHeight} viewBox={`0 0 ${chartWidth} ${chartHeight}`} preserveAspectRatio="xMidYMid meet">
                {/* Grid lines */}
                {[0, 25, 50, 75, 100].map((tick) => (
                    <line
                        key={tick}
                        x1={padding}
                        y1={chartHeight - 20 - ((chartHeight - 40) * tick / 100)}
                        x2={chartWidth - 10}
                        y2={chartHeight - 20 - ((chartHeight - 40) * tick / 100)}
                        stroke={theme.palette.neutralLight}
                        strokeWidth="1"
                        strokeDasharray="4,4"
                    />
                ))}

                {/* Y-axis labels */}
                {[0, 50, 100].map((tick) => (
                    <text
                        key={`y-${tick}`}
                        x={5}
                        y={chartHeight - 16 - ((chartHeight - 40) * tick / 100)}
                        fontSize="10"
                        fill={theme.palette.neutralSecondary}
                    >
                        {Math.round(minValue + (valueRange * tick / 100))}
                    </text>
                ))}

                {/* Area fill */}
                <path
                    d={createAreaPath()}
                    fill={theme.palette.themeLight}
                    opacity="0.4"
                />

                {/* Line */}
                <path
                    d={createSmoothPath()}
                    fill="none"
                    stroke={theme.palette.themePrimary}
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                />

                {/* Data points */}
                {points.map((point) => (
                    <g key={point.label}>
                        <circle
                            cx={point.x}
                            cy={point.y}
                            r="4"
                            fill={theme.palette.white}
                            stroke={theme.palette.themePrimary}
                            strokeWidth="2"
                        />
                        {/* X-axis label */}
                        <text
                            x={point.x}
                            y={chartHeight - 5}
                            textAnchor="middle"
                            fontSize="11"
                            fontWeight="500"
                            fill={theme.palette.neutralSecondary}
                        >
                            {point.label}
                        </text>
                    </g>
                ))}
            </svg>

            {/* Legend - matching DonutChart style */}
            <div style={{
                display: 'flex',
                justifyContent: 'center',
                gap: '16px',
                flexWrap: 'wrap',
                padding: '0 8px',
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <div style={{
                        width: '14px',
                        height: '3px',
                        backgroundColor: theme.palette.themePrimary,
                        borderRadius: '2px'
                    }} />
                    <div style={{ display: 'flex', flexDirection: 'column' }}>
                        <span style={{ fontSize: '12px', fontWeight: 500, color: theme.palette.neutralPrimary }}>
                            Monthly Trend
                        </span>
                        <span style={{ fontSize: '10px', color: theme.palette.neutralSecondary }}>
                            Activity over time
                        </span>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default DashboardLineChart;
