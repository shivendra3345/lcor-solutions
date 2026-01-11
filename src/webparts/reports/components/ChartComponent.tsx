import * as React from 'react';
import {
    Chart as ChartJS,
    CategoryScale,
    LinearScale,
    PointElement,
    LineElement,
    BarElement,
    ArcElement,
    Tooltip,
    Legend,
    ChartOptions,
    TooltipItem
} from 'chart.js';
import { Bar, Line, Pie, Doughnut } from 'react-chartjs-2';
import { CSVData, CSVRow } from '../services/CSVDataService';
import styles from './ChartComponent.module.scss';

// Register ChartJS components
ChartJS.register(
    CategoryScale,
    LinearScale,
    PointElement,
    LineElement,
    BarElement,
    ArcElement,
    Tooltip,
    Legend
);

export interface IChartComponentProps {
    data: CSVData;
    xAxis: string;
    yAxis: string;
    chartType: 'bar' | 'line' | 'pie' | 'doughnut';
    isDarkTheme?: boolean;
    chartTitle?: string;
    hideAxisNames?: boolean;
}

/**
 * Chart Component that visualizes CSV data using Chart.js
 */
export default class ChartComponent extends React.Component<IChartComponentProps> {
    private readonly colors = [
        '#0078d4', // Microsoft Blue
        '#107c10', // Green
        '#d83b01', // Orange Red
        '#8661c5', // Purple
        '#00b7c3', // Cyan
        '#f50f0f', // Red
        '#ffb900', // Gold
        '#00bcf2'  // Light Blue
    ];

    private prepareChartData() {
        const { data, xAxis, yAxis } = this.props;
        const labels: string[] = [];
        const values: number[] = [];

        // Extract values for the selected axes
        data.rows.forEach((row: CSVRow) => {
            const xValue = row[xAxis];
            const yValue = row[yAxis];

            if (xValue !== undefined && yValue !== undefined) {
                labels.push(String(xValue));
                values.push(typeof yValue === 'number' ? yValue : parseFloat(String(yValue)) || 0);
            }
        });

        return { labels, values };
    }

    private getChartOptions(): ChartOptions<any> {
        const { isDarkTheme } = this.props;
        const hideAxis = !!this.props.hideAxisNames;
        const textColor = isDarkTheme ? '#ffffff' : '#333333';

        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: {
                        color: textColor,
                        font: {
                            size: 12,
                            weight: 'bold'
                        }
                    }
                },
                tooltip: {
                    backgroundColor: isDarkTheme ? 'rgba(0, 0, 0, 0.8)' : 'rgba(255, 255, 255, 0.9)',
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: isDarkTheme ? '#666666' : '#cccccc',
                    borderWidth: 1,
                    callbacks: {
                        label: (context: TooltipItem<any>) => {
                            const value = context.parsed.y || context.parsed;
                            return `${this.props.yAxis}: ${typeof value === 'number' ? value.toFixed(2) : value}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: textColor
                    },
                    grid: {
                        color: isDarkTheme ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)'
                    },
                    title: {
                        display: !hideAxis,
                        text: this.props.xAxis,
                        color: textColor,
                        font: {
                            size: 14,
                            weight: 'bold'
                        }
                    }
                },
                y: {
                    ticks: {
                        color: textColor
                    },
                    grid: {
                        color: isDarkTheme ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)'
                    },
                    title: {
                        display: !hideAxis,
                        text: this.props.yAxis,
                        color: textColor,
                        font: {
                            size: 14,
                            weight: 'bold'
                        }
                    }
                }
            }
        } as ChartOptions<any>;
    }

    private getPieChartOptions(): ChartOptions<any> {
        const { isDarkTheme } = this.props;
        const textColor = isDarkTheme ? '#ffffff' : '#333333';

        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: textColor,
                        font: {
                            size: 12
                        }
                    }
                },
                tooltip: {
                    backgroundColor: isDarkTheme ? 'rgba(0, 0, 0, 0.8)' : 'rgba(255, 255, 255, 0.9)',
                    titleColor: textColor,
                    bodyColor: textColor,
                    callbacks: {
                        label: (context: TooltipItem<any>) => {
                            const total = context.dataset.data.reduce((a: number, b: number) => a + b, 0);
                            const percentage = ((context.parsed as number / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed} (${percentage}%)`;
                        }
                    }
                }
            }
        } as ChartOptions<any>;
    }

    public render(): React.ReactElement<IChartComponentProps> {
        const { xAxis, yAxis, chartType } = this.props;
        const { labels, values } = this.prepareChartData();

        if (labels.length === 0) {
            return <div className={styles.error}>No data available for the selected columns</div>;
        }

        const baseChartData = {
            labels,
            datasets: [
                {
                    label: yAxis,
                    data: values,
                    backgroundColor: this.colors.slice(0, Math.min(this.colors.length, labels.length)),
                    borderColor: this.colors[0],
                    borderWidth: 2
                }
            ]
        };

        const commonOptions = this.getChartOptions();
        const pieOptions = this.getPieChartOptions();

        return (
            <div className={styles.chartComponent}>
                <div className={styles.chartTitle}>
                    {this.props.chartTitle ? this.props.chartTitle : `${yAxis} by ${xAxis}`}
                </div>
                <div className={styles.chartWrapper}>
                    {chartType === 'bar' && (
                        <Bar data={baseChartData} options={commonOptions} />
                    )}
                    {chartType === 'line' && (
                        <Line data={baseChartData} options={commonOptions} />
                    )}
                    {chartType === 'pie' && (
                        <Pie data={baseChartData} options={pieOptions} />
                    )}
                    {chartType === 'doughnut' && (
                        <Doughnut data={baseChartData} options={pieOptions} />
                    )}
                </div>
            </div>
        );
    }
}
