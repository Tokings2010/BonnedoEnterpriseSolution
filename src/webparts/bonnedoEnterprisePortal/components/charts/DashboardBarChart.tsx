import * as React from 'react';
import { getTheme } from '@fluentui/react';

interface IChartData {
    label: string;
    value: number;
    color: string;
}

interface IDashboardBarChartProps {
    data: IChartData[];
    height?: number;
}

const DashboardBarChart: React.FC<IDashboardBarChartProps> = ({
    data = [],
    height = 200,
}) => {
    const theme = getTheme();

    if (!data || data.length === 0) {
        return <div style={{ height, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#666' }}>No data</div>;
    }

    // Fixed dimensions for proper aspect ratio
    const chartWidth = 280;
    const chartHeight = 140;
    const maxValue = Math.max(...data.map(d => d.value), 1);
    const totalValue = data.reduce((sum, d) => sum + d.value, 0);

    return (
        <div style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            gap: '12px',
            overflow: 'hidden',
            width: '100%',
        }}>
            {/* Bar Chart */}
            <svg width="100%" height={chartHeight} viewBox={`0 0 ${chartWidth} ${chartHeight}`} preserveAspectRatio="xMidYMid meet">
                {/* Grid lines */}
                {[0, 25, 50, 75, 100].map((tick) => (
                    <line
                        key={tick}
                        x1="30"
                        y1={chartHeight - (chartHeight * tick / 100)}
                        x2={chartWidth - 10}
                        y2={chartHeight - (chartHeight * tick / 100)}
                        stroke={theme.palette.neutralLight}
                        strokeWidth="1"
                        strokeDasharray="4,4"
                    />
                ))}

                {/* Bars */}
                {data.map((item, index) => {
                    const barHeight = (item.value / maxValue) * (chartHeight - 30);
                    const barWidth = (chartWidth - 40) / data.length - 15;
                    const x = 35 + index * (barWidth + 15);
                    const y = chartHeight - 25 - barHeight;

                    return (
                        <g key={item.label}>
                            <rect
                                x={x}
                                y={y}
                                width={barWidth}
                                height={barHeight}
                                fill={item.color}
                                rx="4"
                                ry="4"
                            />
                            {/* Label below */}
                            <text
                                x={x + barWidth / 2}
                                y={chartHeight - 8}
                                textAnchor="middle"
                                fontSize="11"
                                fontWeight="500"
                                fill={theme.palette.neutralSecondary}
                            >
                                {item.label}
                            </text>
                            {/* Value on top */}
                            {barHeight > 20 && (
                                <text
                                    x={x + barWidth / 2}
                                    y={y - 6}
                                    textAnchor="middle"
                                    fontSize="11"
                                    fontWeight="600"
                                    fill={theme.palette.neutralPrimary}
                                >
                                    {item.value}
                                </text>
                            )}
                        </g>
                    );
                })}
            </svg>

            {/* Legend - matching DonutChart style */}
            <div style={{
                display: 'flex',
                justifyContent: 'center',
                gap: '16px',
                flexWrap: 'wrap',
                padding: '0 8px',
            }}>
                {data.map((item) => {
                    const percentage = totalValue > 0 ? Math.round((item.value / totalValue) * 100) : 0;
                    return (
                        <div key={item.label} style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <div style={{
                                width: '14px',
                                height: '14px',
                                backgroundColor: item.color,
                                borderRadius: '3px'
                            }} />
                            <div style={{ display: 'flex', flexDirection: 'column' }}>
                                <span style={{ fontSize: '12px', fontWeight: 500, color: theme.palette.neutralPrimary }}>
                                    {item.label}
                                </span>
                                <span style={{ fontSize: '10px', color: theme.palette.neutralSecondary }}>
                                    {item.value} ({percentage}%)
                                </span>
                            </div>
                        </div>
                    );
                })}
            </div>
        </div>
    );
};

export default DashboardBarChart;
