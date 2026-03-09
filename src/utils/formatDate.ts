// ============================================================================
// DWx Policy Manager — Shared Date Formatting Utilities
// Ensures consistent date/time formatting across all components
// ============================================================================

/**
 * Format a date string or Date object to a consistent short format.
 * Output: "5 Jun 2025"
 */
export function formatDate(value: string | Date | undefined | null): string {
  if (!value) return '';
  const date = typeof value === 'string' ? new Date(value) : value;
  if (isNaN(date.getTime())) return '';
  return date.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
  });
}

/**
 * Format a date string or Date object to include time.
 * Output: "5 Jun 2025, 14:32"
 */
export function formatDateTime(value: string | Date | undefined | null): string {
  if (!value) return '';
  const date = typeof value === 'string' ? new Date(value) : value;
  if (isNaN(date.getTime())) return '';
  return date.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

/**
 * Format a date as relative time (e.g., "2 hours ago", "3 days ago").
 */
export function formatRelativeDate(value: string | Date | undefined | null): string {
  if (!value) return '';
  const date = typeof value === 'string' ? new Date(value) : value;
  if (isNaN(date.getTime())) return '';

  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  if (diffMins < 1) return 'Just now';
  if (diffMins < 60) return `${diffMins}m ago`;
  if (diffHours < 24) return `${diffHours}h ago`;
  if (diffDays < 7) return `${diffDays}d ago`;
  return formatDate(date);
}
