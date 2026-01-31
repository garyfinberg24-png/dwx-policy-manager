// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { IStatCard, IBreadcrumbItem, INavItem, IQuickLink } from '../JmlAppHeader/IJmlAppHeaderProps';
import { UserRole } from '../../services/RoleDetectionService';
import { PolicyManagerRole } from '../../services/PolicyRoleService';
import { IFooterLinkGroup } from '../JmlAppFooter/IJmlAppFooterProps';

/**
 * Props for the JML App Layout wrapper component
 * Combines header and footer configuration for a complete page layout
 */
export interface IJmlAppLayoutProps {
  /** SPFx WebPart context */
  context: WebPartContext;
  /** PnPjs SPFI instance for SharePoint operations */
  sp?: SPFI;

  // Page Header Configuration
  /** Page title displayed in header */
  pageTitle: string;
  /** Page description displayed below title */
  pageDescription?: string;
  /** Fluent UI icon name for the page */
  pageIcon?: string;
  /** Stat cards displayed in page header */
  stats?: IStatCard[];
  /** Breadcrumb navigation items */
  breadcrumbs?: IBreadcrumbItem[];

  // Navigation Configuration
  /** Key of the currently active nav item */
  activeNavKey?: string;
  /** Custom navigation items (overrides default) */
  navItems?: INavItem[];
  /** Quick links for circular shortcuts */
  quickLinks?: IQuickLink[];
  /** Whether to show quick links */
  showQuickLinks?: boolean;

  // User/Search/Notifications
  /** Whether to show the search button */
  showSearch?: boolean;
  /** Whether to show the notifications button */
  showNotifications?: boolean;
  /** Notification count badge */
  notificationCount?: number;
  /** Callback when search is clicked */
  onSearchClick?: () => void;
  /** Callback when notifications is clicked */
  onNotificationsClick?: () => void;
  /** Whether to show the settings button */
  showSettings?: boolean;
  /** Callback when settings is clicked */
  onSettingsClick?: () => void;
  /** Current user's role for role-based filtering. Uses UserRole enum. */
  userRole?: UserRole;
  /** Available roles for the current user (shows role badge if >1) */
  availableRoles?: UserRole[];
  /** Callback when role is changed via header badge dropdown */
  onRoleChange?: (role: UserRole) => void;
  /** Policy Manager role for navigation and visibility filtering */
  policyManagerRole?: PolicyManagerRole;

  // Footer Configuration
  /** Application version string */
  version?: string;
  /** Support link URL */
  supportUrl?: string;
  /** Support link text */
  supportText?: string;
  /** Footer link groups */
  footerLinkGroups?: IFooterLinkGroup[];
  /** Whether to show compact footer */
  compactFooter?: boolean;
  /** Organization name for copyright */
  organizationName?: string;

  // Layout Options
  /** Whether to show the header */
  showHeader?: boolean;
  /** Whether to show the footer */
  showFooter?: boolean;
  /** Maximum content width (default: 1400px) */
  maxContentWidth?: string;
  /** Content padding */
  contentPadding?: string;
  /** Children to render in content area */
  children?: React.ReactNode;

}
