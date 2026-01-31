// @ts-nocheck
/**
 * DwxAppHeader - Policy Manager branded header component
 * Uses the PolicyManagerHeader internally for consistent DWx branding
 *
 * Note: File retains Jml naming for import compatibility
 */
import * as React from 'react';
import { IJmlAppHeaderProps } from './IJmlAppHeaderProps';
import { PolicyManagerHeader } from '../PolicyManagerHeader';

export const DwxAppHeader: React.FC<IJmlAppHeaderProps> = ({
  context,
  sp,
  pageTitle,
  pageDescription,
  pageIcon,
  stats,
  breadcrumbs,
  activeNavKey,
  navItems,
  quickLinks,
  showQuickLinks = true,
  showSearch = true,
  showNotifications = true,
  notificationCount = 0,
  onSearchClick,
  onNotificationsClick,
  showSettings = true,
  onSettingsClick,
  userRole,
  availableRoles,
  onRoleChange,
  policyRole
}) => {
  // Get current user info from context if available
  const userName = context?.pageContext?.user?.displayName || 'User';
  const userEmail = context?.pageContext?.user?.email || '';

  // Generate initials from name
  const userInitials = userName
    .split(' ')
    .map((n: string) => n[0])
    .join('')
    .slice(0, 2)
    .toUpperCase();

  // Convert quickLinks to the format expected by PolicyManagerHeader
  const quickActions = showQuickLinks && quickLinks
    ? quickLinks.map((link) => ({
        text: link.text,
        icon: link.icon,
        onClick: link.onClick,
        primary: link.primary
      }))
    : [];

  return (
    <PolicyManagerHeader
      sp={sp}
      userName={userName}
      userEmail={userEmail}
      userInitials={userInitials}
      navItems={navItems}
      activeNavKey={activeNavKey}
      quickActions={quickActions}
      showSearch={showSearch}
      searchPlaceholder="Search policies..."
      onSearch={onSearchClick}
      showNotifications={showNotifications}
      notificationCount={notificationCount}
      onNotificationsClick={onNotificationsClick}
      showSettings={showSettings}
      onSettingsClick={onSettingsClick}
      breadcrumbs={breadcrumbs}
      pageTitle={pageTitle}
      pageDescription={pageDescription}
      pageIcon={pageIcon}
      pageStats={stats}
      showPageHeader={false}
      policyRole={policyRole}
    />
  );
};

// Export with both names for compatibility
export default DwxAppHeader;
// Legacy alias for backward compatibility
export const JmlAppHeader = DwxAppHeader;
export type { IJmlAppHeaderProps };
