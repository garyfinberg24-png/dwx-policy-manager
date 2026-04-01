// @ts-nocheck
import * as React from 'react';
import { EventChannel } from '../../../../models/IEventViewer';
import { ChannelColors } from '../EventViewerStyles';

export interface IChannelBadgeProps {
  channel: EventChannel;
}

export const ChannelBadge: React.FC<IChannelBadgeProps> = ({ channel }) => {
  const colors = ChannelColors[channel] || ChannelColors['System'];

  return (
    <span style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: 4,
      padding: '2px 8px',
      borderRadius: 4,
      fontSize: 11,
      fontWeight: 600,
      background: colors.bg,
      color: colors.text,
      whiteSpace: 'nowrap',
    }}>
      {channel}
    </span>
  );
};
