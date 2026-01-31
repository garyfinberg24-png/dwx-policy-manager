// @ts-nocheck
import * as React from 'react';
import styles from './PolicyManagerHeader.module.scss';
import { SPFI } from '@pnp/sp';
import { PolicyManagerRole, filterNavForRole, getHeaderVisibility } from '../../services/PolicyRoleService';
import { PolicyRequestWizard } from './PolicyRequestWizard';

export interface INavItem {
  key: string;
  text: string;
  icon?: React.ReactNode;
  href?: string;
  onClick?: () => void;
  badge?: number;
  badgeColor?: 'red' | 'green' | 'orange';
  hasDropdown?: boolean;
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

export interface INotificationItem {
  id: number;
  title: string;
  message: string;
  type: 'task' | 'approval' | 'reminder' | 'alert';
  priority: 'high' | 'medium' | 'low';
  time: string;
  isRead: boolean;
}

export interface IPolicyManagerHeaderProps {
  /** Current user's display name */
  userName?: string;
  /** Current user's email */
  userEmail?: string;
  /** Current user's initials for avatar fallback */
  userInitials?: string;
  /** User's profile photo URL */
  userPhotoUrl?: string;
  /** Navigation items */
  navItems?: INavItem[];
  /** Currently active navigation key */
  activeNavKey?: string;
  /** Quick action buttons */
  quickActions?: Array<{
    text: string;
    icon?: React.ReactNode;
    onClick?: () => void;
    primary?: boolean;
  }>;
  /** Show search bar */
  showSearch?: boolean;
  /** Search placeholder text */
  searchPlaceholder?: string;
  /** Search callback */
  onSearch?: (query: string) => void;
  /** Show notifications button */
  showNotifications?: boolean;
  /** Notification count */
  notificationCount?: number;
  /** Notification items */
  notifications?: INotificationItem[];
  /** Notifications callback */
  onNotificationsClick?: () => void;
  /** Show settings button */
  showSettings?: boolean;
  /** Settings callback */
  onSettingsClick?: () => void;
  /** Show help button */
  showHelp?: boolean;
  /** Help callback */
  onHelpClick?: () => void;
  /** Breadcrumbs */
  breadcrumbs?: IBreadcrumb[];
  /** Page title */
  pageTitle?: string;
  /** Page description */
  pageDescription?: string;
  /** Page icon */
  pageIcon?: React.ReactNode;
  /** Page stats */
  pageStats?: IPageStat[];
  /** Show page header section */
  showPageHeader?: boolean;
  /** Login time for profile panel */
  loginTime?: string;
  /** Policy Manager role for nav filtering */
  policyRole?: PolicyManagerRole;
  /** PnPjs SPFI instance for SharePoint operations (wizard submit) */
  sp?: SPFI;
}

// Icon components for nav items
const NavIcons = {
  create: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 4v16m8-8H4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
    </svg>
  ),
  browse: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M4 6h16M4 10h16M4 14h16M4 18h16" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
    </svg>
  ),
  authored: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  approvals: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  delegations: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <circle cx="9" cy="7" r="4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M23 21v-2a4 4 0 00-3-3.87m-4-12a4 4 0 010 7.75" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  analytics: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M18 20V10M12 20V4M6 20v-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  admin: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 15a3 3 0 100-6 3 3 0 000 6z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  packs: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  quiz: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  details: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M14 2v6h6M16 13H8M16 17H8M10 9H8" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  distribution: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M22 2L11 13" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M22 2L15 22l-4-9-9-4 20-7z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  manager: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M16 21v-2a4 4 0 00-4-4H6a4 4 0 00-4 4v2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <circle cx="9" cy="7" r="4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M22 21v-2a4 4 0 00-3-3.87" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M16 3.13a4 4 0 010 7.75" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  dropdown: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M6 9l6 6 6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  )
};

// Notification type icon mapping
const getNotificationIcon = (type: string): string => {
  switch (type) {
    case 'approval': return 'CheckMark';
    case 'reminder': return 'Clock';
    case 'alert': return 'Warning';
    case 'task': return 'TaskSolid';
    default: return 'Info';
  }
};

// Notification type color mapping
const getNotificationColor = (type: string): string => {
  switch (type) {
    case 'approval': return '#107c10';
    case 'reminder': return '#0078d4';
    case 'alert': return '#d83b01';
    case 'task': return '#8764b8';
    default: return '#605e5c';
  }
};

// Default mock notifications
const defaultNotifications: INotificationItem[] = [
  { id: 1, title: 'Policy Approval Required', message: 'Data Protection Policy v3.2 needs your review', type: 'approval', priority: 'high', time: '5m ago', isRead: false },
  { id: 2, title: 'Policy Expiring Soon', message: 'IT Security Policy expires in 14 days', type: 'reminder', priority: 'medium', time: '1h ago', isRead: false },
  { id: 3, title: 'Compliance Alert', message: 'GDPR training acknowledgement overdue', type: 'alert', priority: 'high', time: '3h ago', isRead: false },
  { id: 4, title: 'New Policy Published', message: 'Remote Work Policy v2.0 is now live', type: 'task', priority: 'low', time: '1d ago', isRead: true },
];

/**
 * Policy Manager Header/NavBar Component
 * Based on DWx Brand Guide - Forest Teal theme
 * Nav bar styled like Contract Manager
 * Includes: Profile dropdown, Notifications dropdown, Help & Search navigation
 */
export const PolicyManagerHeader: React.FC<IPolicyManagerHeaderProps> = ({
  userName = 'User',
  userEmail = '',
  userInitials,
  userPhotoUrl,
  navItems = [],
  activeNavKey,
  quickActions = [],
  showSearch = true,
  searchPlaceholder = 'Search policies...',
  onSearch,
  showNotifications = true,
  notificationCount = 0,
  notifications,
  onNotificationsClick,
  showSettings = true,
  onSettingsClick,
  showHelp = true,
  onHelpClick,
  breadcrumbs,
  pageTitle,
  pageDescription,
  pageIcon,
  pageStats,
  showPageHeader = false,
  loginTime,
  policyRole,
  sp
}) => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [showProfileDropdown, setShowProfileDropdown] = React.useState(false);
  const [showNotificationDropdown, setShowNotificationDropdown] = React.useState(false);
  const [showRecentlyViewedDropdown, setShowRecentlyViewedDropdown] = React.useState(false);

  // Request Policy Wizard — extracted to PolicyRequestWizard.tsx
  const [showRequestWizard, setShowRequestWizard] = React.useState(false);

  // Refs for click-outside detection
  const profileRef = React.useRef<HTMLDivElement>(null);
  const notificationRef = React.useRef<HTMLDivElement>(null);
  const recentlyViewedRef = React.useRef<HTMLDivElement>(null);

  // Mock recently viewed policies
  const recentlyViewedPolicies = [
    { id: 1, title: 'Data Protection Policy v3.2', category: 'Data Privacy', time: '10m ago', status: 'Published' },
    { id: 2, title: 'IT Security Policy', category: 'IT & Security', time: '1h ago', status: 'Published' },
    { id: 3, title: 'Remote Work Policy v2.0', category: 'HR Policies', time: '2h ago', status: 'Published' },
    { id: 4, title: 'GDPR Compliance Framework', category: 'Compliance', time: '3h ago', status: 'Published' },
    { id: 5, title: 'Acceptable Use Policy', category: 'IT & Security', time: '1d ago', status: 'In Review' }
  ];

  // Click-outside handler to close dropdowns
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (profileRef.current && !profileRef.current.contains(event.target as Node)) {
        setShowProfileDropdown(false);
      }
      if (notificationRef.current && !notificationRef.current.contains(event.target as Node)) {
        setShowNotificationDropdown(false);
      }
      if (recentlyViewedRef.current && !recentlyViewedRef.current.contains(event.target as Node)) {
        setShowRecentlyViewedDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Generate user initials if not provided
  const displayInitials = userInitials || userName
    .split(' ')
    .map(n => n[0])
    .join('')
    .slice(0, 2)
    .toUpperCase();

  const handleSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      if (onSearch) {
        onSearch(searchQuery);
      }
      // Navigate to search page with query parameter
      if (searchQuery.trim()) {
        window.location.href = `/sites/PolicyManager/SitePages/PolicySearch.aspx?q=${encodeURIComponent(searchQuery.trim())}`;
      }
    }
  };

  // Use provided notifications or defaults
  const displayNotifications = notifications || defaultNotifications;
  const unreadCount = notificationCount || displayNotifications.filter(n => !n.isRead).length;

  // Format login time
  const displayLoginTime = loginTime || new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  // Default nav items — ordered by workflow: consume → author → manage → analyse
  // Role filtering applied via PolicyRoleService.filterNavForRole()
  const defaultNavItems: INavItem[] = [
    { key: 'my-policies', text: 'My Policies', icon: NavIcons.authored, href: '/sites/PolicyManager/SitePages/MyPolicies.aspx' },
    { key: 'browse', text: 'Policy Hub', icon: NavIcons.browse, href: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
    { key: 'create', text: 'Policy Builder', icon: NavIcons.create, href: '/sites/PolicyManager/SitePages/PolicyBuilder.aspx' },
    { key: 'author', text: 'Policy Author', icon: NavIcons.authored, href: '/sites/PolicyManager/SitePages/PolicyAuthor.aspx' },
    { key: 'packs', text: 'Policy Packs', icon: NavIcons.packs, href: '/sites/PolicyManager/SitePages/PolicyPacks.aspx' },
    { key: 'distribution', text: 'Distribution', icon: NavIcons.distribution, href: '/sites/PolicyManager/SitePages/PolicyDistribution.aspx' },
    { key: 'manager', text: 'Policy Manager', icon: NavIcons.manager, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx' },
    { key: 'analytics', text: 'Analytics', icon: NavIcons.analytics, href: '/sites/PolicyManager/SitePages/PolicyAnalytics.aspx' },
    { key: 'quiz', text: 'Quiz Builder', icon: NavIcons.quiz, href: '/sites/PolicyManager/SitePages/QuizBuilder.aspx' },
  ];

  const allNavItems = navItems.length > 0 ? navItems : defaultNavItems;
  const displayNavItems = policyRole ? filterNavForRole(allNavItems, policyRole) : allNavItems;

  // Override header visibility based on role
  const roleVisibility = policyRole ? getHeaderVisibility(policyRole) : null;

  // Get badge class based on color
  const getBadgeClass = (color?: string) => {
    switch (color) {
      case 'green': return styles.navBadgeGreen;
      case 'orange': return styles.navBadgeOrange;
      default: return '';
    }
  };

  // Handle search icon click — navigate to search page
  const handleSearchClick = () => {
    window.location.href = '/sites/PolicyManager/SitePages/PolicySearch.aspx';
  };

  // Handle help icon click — navigate to help page
  const handleHelpClick = () => {
    if (onHelpClick) {
      onHelpClick();
    } else {
      window.location.href = '/sites/PolicyManager/SitePages/PolicyHelp.aspx';
    }
  };

  // Handle notification bell click
  const handleNotificationClick = () => {
    setShowNotificationDropdown(!showNotificationDropdown);
    setShowProfileDropdown(false);
  };

  // Handle profile avatar click
  const handleProfileClick = () => {
    setShowProfileDropdown(!showProfileDropdown);
    setShowNotificationDropdown(false);
  };

  return (
    <>
    <header className={styles.headerContainer}>
      {/* Top Bar - Dark teal gradient */}
      <div className={styles.topBar}>
        {/* Logo Section */}
        <div className={styles.logoSection}>
          <div className={styles.logoIcon}>
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path
                d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </div>
          <div className={styles.logoText}>
            <span className={styles.appName}>Policy Manager</span>
            <span className={styles.appSubtitle}>Policy Governance & Compliance</span>
          </div>
        </div>

        {/* Search Section */}
        {showSearch && (
          <div className={styles.searchSection}>
            <div className={styles.searchInput}>
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path
                  d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
              <input
                type="text"
                placeholder={searchPlaceholder}
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                onKeyDown={handleSearchKeyDown}
              />
            </div>
          </div>
        )}

        {/* Actions Section */}
        <div className={styles.actionsSection}>
          {/* Mobile Menu Button */}
          <button className={styles.mobileMenuButton} type="button">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path
                d="M4 6h16M4 12h16M4 18h16"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>

          {/* Quick Action Buttons */}
          <div className={styles.quickActionButtons}>
            <button
              className={`${styles.quickActionBtn} ${styles.quickActionBtnPrimary}`}
              type="button"
              title="New Policy"
              aria-label="New Policy"
              onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyBuilder.aspx'}
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
              New Policy
            </button>
            <button
              className={styles.quickActionBtn}
              type="button"
              title="Request Policy"
              aria-label="Request Policy"
              onClick={() => setShowRequestWizard(true)}
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M5 12h14M12 5l7 7-7 7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
              Request Policy
            </button>
            <div className={styles.dropdownContainer} ref={recentlyViewedRef} style={{ display: 'inline-flex' }}>
              <button
                className={styles.quickActionBtn}
                type="button"
                title="Recently Viewed"
                aria-label="Recently Viewed"
                onClick={() => {
                  setShowRecentlyViewedDropdown(!showRecentlyViewedDropdown);
                  setShowProfileDropdown(false);
                  setShowNotificationDropdown(false);
                }}
              >
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2"/>
                  <path d="M12 6v6l4 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
                Recently Viewed
              </button>

              {showRecentlyViewedDropdown && (
                <div className={styles.dropdownPanel} style={{ left: 0, right: 'auto', minWidth: '340px' }}>
                  <div className={styles.dropdownArrow} style={{ left: '40px', right: 'auto' }} />
                  <div className={styles.dropdownPanelHeader}>
                    <span className={styles.dropdownPanelTitle}>Recently Viewed</span>
                    <span className={styles.dropdownPanelBadge}>{recentlyViewedPolicies.length} policies</span>
                  </div>
                  <div className={styles.dropdownPanelBody}>
                    {recentlyViewedPolicies.map(policy => (
                      <a
                        key={policy.id}
                        href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.id}`}
                        className={styles.notificationItem}
                        style={{ textDecoration: 'none', color: 'inherit' }}
                        onClick={() => setShowRecentlyViewedDropdown(false)}
                      >
                        <div
                          className={styles.notificationItemIcon}
                          style={{ background: '#0d9488' }}
                        >
                          <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                            <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                            <path d="M14 2v6h6M16 13H8M16 17H8M10 9H8" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          </svg>
                        </div>
                        <div className={styles.notificationItemContent}>
                          <span className={styles.notificationItemTitle}>{policy.title}</span>
                          <span className={styles.notificationItemMessage}>{policy.category}</span>
                        </div>
                        <span className={styles.notificationItemTime}>{policy.time}</span>
                      </a>
                    ))}
                  </div>
                  <div className={styles.dropdownPanelFooter}>
                    <a href="/sites/PolicyManager/SitePages/PolicyHub.aspx?view=recent" className={styles.dropdownFooterLink}>
                      View All Recent Policies
                    </a>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Search icon button */}
          <button
            className={styles.actionButton}
            type="button"
            title="Search"
            onClick={handleSearchClick}
            aria-label="Search"
          >
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          </button>

          {/* Help */}
          {showHelp && (
            <button
              className={styles.actionButton}
              onClick={handleHelpClick}
              type="button"
              title="Help Center"
              aria-label="Help Center"
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M9.09 9a3 3 0 015.83 1c0 2-3 3-3 3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M12 17h.01" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </button>
          )}

          {/* Admin Settings (Cog) — visible if explicitly enabled OR if role allows it */}
          {(showSettings || (roleVisibility ? roleVisibility.showSettings : false)) && (
            <button
              className={styles.actionButton}
              onClick={() => {
                if (onSettingsClick) {
                  onSettingsClick();
                } else {
                  window.location.href = '/sites/PolicyManager/SitePages/PolicyAdmin.aspx';
                }
              }}
              type="button"
              title="Policy Administration"
              aria-label="Policy Administration"
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M12 15a3 3 0 100-6 3 3 0 000 6z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z" stroke="currentColor" strokeWidth="2"/>
              </svg>
            </button>
          )}

          {/* Notifications with Dropdown */}
          {showNotifications && (
            <div className={styles.dropdownContainer} ref={notificationRef}>
              <button
                className={styles.actionButton}
                onClick={handleNotificationClick}
                type="button"
                title="Notifications"
                aria-label="Notifications"
              >
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path
                    d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 01-3.46 0"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
                {unreadCount > 0 && (
                  <span className={styles.notificationBadgeCount} style={{ border: 'none', outline: 'none', boxShadow: 'none' }}>
                    {unreadCount > 99 ? '99+' : unreadCount}
                  </span>
                )}
              </button>

              {/* Notifications Dropdown Panel */}
              {showNotificationDropdown && (
                <div className={styles.dropdownPanel}>
                  <div className={styles.dropdownArrow} />
                  <div className={styles.dropdownPanelHeader}>
                    <span className={styles.dropdownPanelTitle}>Notifications</span>
                    <span className={styles.dropdownPanelBadge}>{unreadCount} new</span>
                  </div>
                  <div className={styles.dropdownPanelBody}>
                    {displayNotifications.length === 0 ? (
                      <div className={styles.dropdownEmpty}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 32, height: 32, color: '#c8c6c4' }}>
                          <path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 01-3.46 0" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        <span>No new notifications</span>
                      </div>
                    ) : (
                      displayNotifications.map(notification => (
                        <div
                          key={notification.id}
                          className={`${styles.notificationItem} ${!notification.isRead ? styles.notificationUnread : ''}`}
                        >
                          <div
                            className={styles.notificationItemIcon}
                            style={{ background: getNotificationColor(notification.type) }}
                          >
                            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                              {notification.type === 'approval' && <path d="M20 6L9 17l-5-5" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>}
                              {notification.type === 'reminder' && <><circle cx="12" cy="12" r="10" stroke="white" strokeWidth="2"/><path d="M12 6v6l4 2" stroke="white" strokeWidth="2" strokeLinecap="round"/></>}
                              {notification.type === 'alert' && <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0zM12 9v4M12 17h.01" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>}
                              {notification.type === 'task' && <><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" stroke="white" strokeWidth="2" strokeLinecap="round"/><path d="M9 12l2 2 4-4" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></>}
                            </svg>
                          </div>
                          <div className={styles.notificationItemContent}>
                            <span className={styles.notificationItemTitle}>{notification.title}</span>
                            <span className={styles.notificationItemMessage}>{notification.message}</span>
                          </div>
                          <span className={styles.notificationItemTime}>{notification.time}</span>
                        </div>
                      ))
                    )}
                  </div>
                  <div className={styles.dropdownPanelFooter}>
                    <a href="/sites/PolicyManager/SitePages/PolicyHub.aspx?view=notifications" className={styles.dropdownFooterLink}>
                      View All Notifications
                    </a>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* User Avatar with Profile Dropdown */}
          <div className={styles.dropdownContainer} ref={profileRef}>
            <div
              className={styles.userAvatar}
              title={userName}
              onClick={handleProfileClick}
              role="button"
              tabIndex={0}
              aria-label="User profile"
            >
              {userPhotoUrl ? (
                <img src={userPhotoUrl} alt={userName} />
              ) : (
                displayInitials
              )}
            </div>

            {/* Profile Dropdown Panel */}
            {showProfileDropdown && (
              <div className={styles.dropdownPanel}>
                <div className={styles.dropdownArrow} />
                <div className={styles.profileHeader}>
                  <div className={styles.profileAvatar}>
                    {userPhotoUrl ? (
                      <img src={userPhotoUrl} alt={userName} />
                    ) : (
                      displayInitials
                    )}
                  </div>
                  <div className={styles.profileName}>{userName}</div>
                  {userEmail && <div className={styles.profileEmail}>{userEmail}</div>}
                  <div className={styles.profileLoginTime}>
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 12, height: 12 }}>
                      <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2"/>
                      <path d="M12 6v6l4 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                    </svg>
                    Logged in at {displayLoginTime}
                  </div>
                </div>
                <div className={styles.profileActions}>
                  <a
                    href={`https://delve.office.com`}
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.profileActionItem}
                  >
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                      <path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                      <circle cx="12" cy="7" r="4" stroke="currentColor" strokeWidth="2"/>
                    </svg>
                    View Profile
                  </a>
                  <a
                    href="https://myaccount.microsoft.com"
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.profileActionItem}
                  >
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                      <path d="M12 15a3 3 0 100-6 3 3 0 000 6z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                      <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z" stroke="currentColor" strokeWidth="2"/>
                    </svg>
                    Account Settings
                  </a>
                </div>
                <div className={styles.profileFooter}>
                  <button
                    className={styles.profileSignOut}
                    onClick={() => {
                      window.location.href = '/_layouts/15/SignOut.aspx';
                    }}
                    type="button"
                  >
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                      <path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                    Sign Out
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Navigation Bar - White background */}
      <nav className={styles.navBar}>
        <div className={styles.navItems}>
          {displayNavItems.map((item) => (
            <a
              key={item.key}
              href={item.href}
              className={`${styles.navItem} ${activeNavKey === item.key ? styles.active : ''}`}
              onClick={item.onClick}
            >
              {item.icon}
              {item.text}
              {item.badge !== undefined && item.badge > 0 && (
                <span className={`${styles.navBadge} ${getBadgeClass(item.badgeColor)}`}>
                  {item.badge}
                </span>
              )}
              {item.hasDropdown && (
                <span className={styles.navDropdown}>
                  {NavIcons.dropdown}
                </span>
              )}
            </a>
          ))}
        </div>

        {/* Quick Actions */}
        {quickActions.length > 0 && (
          <div className={styles.quickActions}>
            {quickActions.map((action, index) => (
              <button
                key={index}
                className={`${styles.quickActionButton} ${action.primary ? styles.primary : styles.secondary}`}
                onClick={action.onClick}
                type="button"
              >
                {action.icon}
                {action.text}
              </button>
            ))}
          </div>
        )}
      </nav>

      {/* Breadcrumbs Bar — aligned with content area */}
      {breadcrumbs && breadcrumbs.length > 0 && (
        <div className={styles.breadcrumbBarOuter}>
          <div className={styles.breadcrumbBar}>
            {breadcrumbs.map((crumb, index) => {
              const link = crumb.href || crumb.url;
              return (
                <span key={index} className={styles.breadcrumbItem}>
                  {index > 0 && <span className={styles.breadcrumbSeparator}>/</span>}
                  {link ? (
                    <a href={link} className={styles.breadcrumbLink}>{crumb.text}</a>
                  ) : (
                    <span className={styles.breadcrumbCurrent}>{crumb.text}</span>
                  )}
                </span>
              );
            })}
          </div>
        </div>
      )}

      {/* Page Header - Hidden by default */}
      {showPageHeader && pageTitle && (
        <div className={styles.pageHeader}>
          <div className={styles.pageHeaderContent}>
            <div className={styles.pageTitleSection}>
              {pageIcon && <div className={styles.pageIcon}>{pageIcon}</div>}
              <div>
                <h1 className={styles.pageTitle}>{pageTitle}</h1>
                {pageDescription && <p className={styles.pageDescription}>{pageDescription}</p>}
              </div>
            </div>

            {pageStats && pageStats.length > 0 && (
              <div className={styles.pageStats}>
                {pageStats.map((stat, index) => (
                  <div key={index} className={styles.pageStat}>
                    <span className={styles.pageStatValue}>{stat.value}</span>
                    <span className={styles.pageStatLabel}>{stat.label}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </header>

    {/* Request Policy Wizard — extracted component */}
    <PolicyRequestWizard
      isOpen={showRequestWizard}
      onClose={() => setShowRequestWizard(false)}
      sp={sp}
      userName={userName}
      userEmail={userEmail}
    />


    </>
  );
};

export default PolicyManagerHeader;
