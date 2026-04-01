// @ts-nocheck
import * as React from 'react';
import { EventSeverity } from '../../../../models/IEventViewer';
import { SeverityColors } from '../EventViewerStyles';

const SEVERITY_LABELS: Record<number, string> = {
  [EventSeverity.Verbose]: 'Verbose',
  [EventSeverity.Information]: 'Info',
  [EventSeverity.Warning]: 'Warning',
  [EventSeverity.Error]: 'Error',
  [EventSeverity.Critical]: 'Critical',
};

const SEVERITY_ICONS: Record<number, JSX.Element> = {
  [EventSeverity.Warning]: (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
      <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
    </svg>
  ),
  [EventSeverity.Error]: (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
    </svg>
  ),
  [EventSeverity.Critical]: (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
      <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
    </svg>
  ),
  [EventSeverity.Information]: (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>
    </svg>
  ),
};

export interface ISeverityBadgeProps {
  severity: EventSeverity;
  compact?: boolean;
}

export const SeverityBadge: React.FC<ISeverityBadgeProps> = ({ severity, compact }) => {
  const label = SEVERITY_LABELS[severity] || 'Unknown';
  const colors = SeverityColors[label] || SeverityColors['Information'];

  return (
    <span style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: 4,
      padding: compact ? '1px 6px' : '2px 8px',
      borderRadius: 4,
      fontSize: compact ? 10 : 11,
      fontWeight: 600,
      textTransform: 'uppercase',
      letterSpacing: 0.3,
      background: colors.bg,
      color: colors.text,
      whiteSpace: 'nowrap',
    }}>
      {SEVERITY_ICONS[severity]}
      {label}
    </span>
  );
};
