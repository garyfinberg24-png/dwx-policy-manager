// @ts-nocheck
import * as React from 'react';
import { SLOW_REQUEST_THRESHOLD_MS } from '../../../../constants/EventCodes';

export interface IWaterfallBarProps {
  duration: number;
  maxDuration?: number;
}

export const WaterfallBar: React.FC<IWaterfallBarProps> = ({ duration, maxDuration = 5000 }) => {
  const widthPercent = Math.min((duration / maxDuration) * 100, 100);
  const color = duration > SLOW_REQUEST_THRESHOLD_MS
    ? '#dc2626'
    : duration > 1000
    ? '#d97706'
    : '#0d9488';

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
      <div style={{
        height: 4,
        borderRadius: 2,
        background: color,
        width: `${Math.max(widthPercent, 3)}%`,
        minWidth: 2,
      }} />
      <span style={{
        fontFamily: "'Cascadia Code', 'Fira Code', 'Consolas', monospace",
        fontSize: 11,
        color: duration > SLOW_REQUEST_THRESHOLD_MS ? '#dc2626' : '#64748b',
        whiteSpace: 'nowrap',
        fontWeight: duration > SLOW_REQUEST_THRESHOLD_MS ? 600 : 400,
      }}>
        {duration >= 1000 ? `${(duration / 1000).toFixed(1)}s` : `${duration}ms`}
      </span>
    </div>
  );
};
