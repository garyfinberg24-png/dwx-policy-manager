// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PolicyManagerRole } from '../../services/PolicyRoleService';

export interface INavItem {
  key: string;
  text: string;
  icon?: React.ReactNode;
  href?: string;
  onClick?: () => void;
}

export interface IBreadcrumb {
  text: string;
  href?: string;
  url?: string;
}

export interface IPageStat {
  value: string | number;
  label: string;
}

export interface IJmlAppHeaderProps {
  /** SPFx WebPart context */
  context?: WebPartContext;
  /** Page title */
  pageTitle?: string;
  /** Page description */
  pageDescription?: string;
  /** Page icon (React node) */
  pageIcon?: React.ReactNode;
  /** Stats to display in header */
  stats?: IPageStat[];
  /** Breadcrumb navigation */
  breadcrumbs?: IBreadcrumb[];
  /** Currently active nav key */
  activeNavKey?: string;
  /** Navigation items */
  navItems?: INavItem[];
  /** Quick link buttons */
  quickLinks?: Array<{
    text: string;
    icon?: React.ReactNode;
    onClick?: () => void;
    primary?: boolean;
  }>;
  /** Show quick links */
  showQuickLinks?: boolean;
  /** Show search bar */
  showSearch?: boolean;
  /** Show notifications */
  showNotifications?: boolean;
  /** Notification count */
  notificationCount?: number;
  /** Search callback */
  onSearchClick?: (query: string) => void;
  /** Notifications callback */
  onNotificationsClick?: () => void;
  /** Show settings button */
  showSettings?: boolean;
  /** Settings callback */
  onSettingsClick?: () => void;
  /** Current user role */
  userRole?: string;
  /** Available roles for switching */
  availableRoles?: string[];
  /** Role change callback */
  onRoleChange?: (role: string) => void;
  /** Policy Manager role for nav/visibility filtering */
  policyRole?: PolicyManagerRole;
}

// Re-export types for backward compatibility
export type { INavItem, IBreadcrumb, IPageStat };
