import * as React from 'react';
import { Icon, Text, PrimaryButton } from '@fluentui/react';

export interface IEmptyStateProps {
  icon?: string;
  title: string;
  message: string;
  actionText?: string;
  actionHref?: string;
  onAction?: () => void;
}

/**
 * Shared empty state component — use when a list/table has zero items.
 * Provides consistent messaging across all webparts.
 */
export const EmptyState: React.FC<IEmptyStateProps> = ({
  icon = 'PageData',
  title,
  message,
  actionText,
  actionHref,
  onAction
}) => {
  const handleAction = (): void => {
    if (onAction) {
      onAction();
    } else if (actionHref) {
      window.location.href = actionHref;
    }
  };

  return (
    <div style={{
      textAlign: 'center',
      padding: '48px 24px',
      background: '#fafafa',
      borderRadius: 4,
      border: '1px dashed #e2e8f0'
    }}>
      <Icon iconName={icon} style={{
        fontSize: 48,
        color: '#94a3b8',
        marginBottom: 16,
        display: 'block'
      }} />
      <Text variant="large" style={{
        display: 'block',
        fontWeight: 600,
        color: '#0f172a',
        marginBottom: 8
      }}>
        {title}
      </Text>
      <Text style={{
        display: 'block',
        color: '#64748b',
        maxWidth: 400,
        margin: '0 auto 16px'
      }}>
        {message}
      </Text>
      {actionText && (
        <PrimaryButton
          text={actionText}
          onClick={handleAction}
          styles={{ root: { borderRadius: 4 } }}
        />
      )}
    </div>
  );
};

export default EmptyState;
