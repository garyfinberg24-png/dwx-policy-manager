// @ts-nocheck
import * as React from 'react';
import { HealthStatus } from '../../../../models/IEventViewer';
import { HealthColors } from '../EventViewerStyles';

export interface IHealthIndicatorProps {
  status: HealthStatus;
  size?: number;
}

export const HealthIndicator: React.FC<IHealthIndicatorProps> = ({ status, size = 12 }) => {
  const color = HealthColors[status] || HealthColors['Healthy'];
  return (
    <div style={{
      width: size,
      height: size,
      borderRadius: '50%',
      background: color,
      boxShadow: `0 0 6px ${color}66`,
      flexShrink: 0,
    }} />
  );
};
