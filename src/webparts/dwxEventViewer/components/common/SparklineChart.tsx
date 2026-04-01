// @ts-nocheck
import * as React from 'react';

export interface ISparklineChartProps {
  data: number[];
  color?: string;
  width?: number;
  height?: number;
}

/**
 * Lightweight SVG sparkline — no external charting library.
 * Renders a polyline from an array of numeric values.
 */
export const SparklineChart: React.FC<ISparklineChartProps> = ({
  data,
  color = '#0d9488',
  width = 80,
  height = 24,
}) => {
  if (!data || data.length < 2) {
    return <div style={{ width, height }} />;
  }

  const max = Math.max(...data, 1);
  const stepX = width / (data.length - 1);
  const padding = 2;
  const chartHeight = height - padding * 2;

  const points = data.map((val, i) => {
    const x = i * stepX;
    const y = padding + chartHeight - (val / max) * chartHeight;
    return `${x},${y}`;
  }).join(' ');

  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ display: 'block' }}>
      <polyline
        points={points}
        fill="none"
        stroke={color}
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
};
