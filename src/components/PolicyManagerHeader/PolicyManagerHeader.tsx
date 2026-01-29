// @ts-nocheck
import * as React from 'react';
import styles from './PolicyManagerHeader.module.scss';
import { PolicyManagerRole, filterNavForRole, getHeaderVisibility } from '../../services/PolicyRoleService';

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
  policyRole
}) => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [showProfileDropdown, setShowProfileDropdown] = React.useState(false);
  const [showNotificationDropdown, setShowNotificationDropdown] = React.useState(false);

  // Request Policy Wizard State
  const [showRequestWizard, setShowRequestWizard] = React.useState(false);
  const [wizardStep, setWizardStep] = React.useState(0);
  const [wizardSubmitted, setWizardSubmitted] = React.useState(false);
  const [requestForm, setRequestForm] = React.useState({
    policyTitle: '',
    policyCategory: '',
    policyType: 'New Policy',
    priority: 'Medium',
    targetAudience: '',
    businessJustification: '',
    regulatoryDriver: '',
    desiredEffectiveDate: '',
    readTimeframeDays: '7',
    requiresAcknowledgement: true,
    requiresQuiz: false,
    additionalNotes: '',
    notifyAuthors: true,
    preferredAuthor: ''
  });

  const updateRequestForm = (field: string, value: string | boolean) => {
    setRequestForm(prev => ({ ...prev, [field]: value }));
  };

  const resetWizard = () => {
    setWizardStep(0);
    setWizardSubmitted(false);
    setRequestForm({
      policyTitle: '', policyCategory: '', policyType: 'New Policy', priority: 'Medium',
      targetAudience: '', businessJustification: '', regulatoryDriver: '',
      desiredEffectiveDate: '', readTimeframeDays: '7', requiresAcknowledgement: true,
      requiresQuiz: false, additionalNotes: '', notifyAuthors: true, preferredAuthor: ''
    });
  };

  const WIZARD_STEPS = [
    { title: 'Policy Details', description: 'What policy do you need?' },
    { title: 'Business Case', description: 'Why is this policy needed?' },
    { title: 'Requirements', description: 'Audience, timeline & compliance' },
    { title: 'Review & Submit', description: 'Confirm and submit your request' }
  ];

  // Refs for click-outside detection
  const profileRef = React.useRef<HTMLDivElement>(null);
  const notificationRef = React.useRef<HTMLDivElement>(null);

  // Click-outside handler to close dropdowns
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (profileRef.current && !profileRef.current.contains(event.target as Node)) {
        setShowProfileDropdown(false);
      }
      if (notificationRef.current && !notificationRef.current.contains(event.target as Node)) {
        setShowNotificationDropdown(false);
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
    if (e.key === 'Enter' && onSearch) {
      onSearch(searchQuery);
    }
  };

  // Use provided notifications or defaults
  const displayNotifications = notifications || defaultNotifications;
  const unreadCount = notificationCount || displayNotifications.filter(n => !n.isRead).length;

  // Format login time
  const displayLoginTime = loginTime || new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  // Default nav items for Policy Manager with icons
  const defaultNavItems: INavItem[] = [
    { key: 'create', text: 'Policy Builder', icon: NavIcons.create, href: '/sites/PolicyManager/SitePages/PolicyBuilder.aspx' },
    { key: 'author', text: 'Policy Author', icon: NavIcons.authored, href: '/sites/PolicyManager/SitePages/PolicyAuthor.aspx' },
    { key: 'manager', text: 'Manager View', icon: NavIcons.manager, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx' },
    { key: 'browse', text: 'Browse Policies', icon: NavIcons.browse, href: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
    { key: 'my-policies', text: 'My Policies', icon: NavIcons.authored, href: '/sites/PolicyManager/SitePages/MyPolicies.aspx' },
    { key: 'distribution', text: 'Distribution', icon: NavIcons.distribution, href: '/sites/PolicyManager/SitePages/PolicyDistribution.aspx' },
    { key: 'analytics', text: 'Analytics', icon: NavIcons.analytics, href: '/sites/PolicyManager/SitePages/PolicyAnalytics.aspx' },
    { key: 'details', text: 'Policy Details', icon: NavIcons.details, href: '/sites/PolicyManager/SitePages/PolicyDetails.aspx' },
    { key: 'packs', text: 'Policy Packs', icon: NavIcons.packs, href: '/sites/PolicyManager/SitePages/PolicyPacks.aspx' },
    { key: 'quiz', text: 'Quiz Builder', icon: NavIcons.quiz, href: '/sites/PolicyManager/SitePages/QuizBuilder.aspx' }
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
              title="New Contract"
              aria-label="New Contract"
              onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyBuilder.aspx'}
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
              New Contract
            </button>
            <button
              className={styles.quickActionBtn}
              type="button"
              title="Request Contract"
              aria-label="Request Contract"
              onClick={() => { resetWizard(); setShowRequestWizard(true); }}
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M5 12h14M12 5l7 7-7 7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
              Request Contract
            </button>
            <button
              className={styles.quickActionBtn}
              type="button"
              title="Recently Viewed"
              aria-label="Recently Viewed"
              onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyHub.aspx?view=recent'}
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2"/>
                <path d="M12 6v6l4 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
              Recently Viewed
            </button>
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
                  <span className={styles.notificationBadgeCount}>
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

      {/* Breadcrumbs Bar */}
      {breadcrumbs && breadcrumbs.length > 0 && (
        <div className={styles.breadcrumbBar}>
          {breadcrumbs.map((crumb, index) => (
            <span key={index} className={styles.breadcrumbItem}>
              {index > 0 && <span className={styles.breadcrumbSeparator}>/</span>}
              {crumb.href ? (
                <a href={crumb.href} className={styles.breadcrumbLink}>{crumb.text}</a>
              ) : (
                <span className={styles.breadcrumbCurrent}>{crumb.text}</span>
              )}
            </span>
          ))}
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

    {/* ================================================================ */}
    {/* REQUEST POLICY WIZARD — Full-screen overlay with stepped form     */}
    {/* ================================================================ */}
    {showRequestWizard && (
      <div className={styles.wizardOverlay}>
        <div className={styles.wizardModal}>
          {/* Wizard Header */}
          <div className={styles.wizardHeader}>
            <div className={styles.wizardHeaderLeft}>
              <div className={styles.wizardHeaderIcon}>
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 24, height: 24 }}>
                  <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  <path d="M14 2v6h6M12 18v-6M9 15h6" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              <div>
                <div className={styles.wizardTitle}>Request a New Policy</div>
                <div className={styles.wizardSubtitle}>Submit a request to the Policy Authoring team</div>
              </div>
            </div>
            <button className={styles.wizardCloseBtn} onClick={() => setShowRequestWizard(false)} title="Close">
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 20, height: 20 }}>
                <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </button>
          </div>

          {/* Step Progress */}
          {!wizardSubmitted && (
            <div className={styles.wizardStepper}>
              {WIZARD_STEPS.map((step, index) => (
                <div key={index} className={styles.wizardStepItem}>
                  <div className={`${styles.wizardStepCircle} ${index < wizardStep ? styles.wizardStepCompleted : ''} ${index === wizardStep ? styles.wizardStepActive : ''}`}>
                    {index < wizardStep ? (
                      <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                        <path d="M20 6L9 17l-5-5" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                    ) : (
                      <span>{index + 1}</span>
                    )}
                  </div>
                  <div className={styles.wizardStepLabel}>
                    <div className={styles.wizardStepTitle}>{step.title}</div>
                    <div className={styles.wizardStepDesc}>{step.description}</div>
                  </div>
                  {index < WIZARD_STEPS.length - 1 && <div className={`${styles.wizardStepConnector} ${index < wizardStep ? styles.wizardStepConnectorDone : ''}`} />}
                </div>
              ))}
            </div>
          )}

          {/* Wizard Body */}
          <div className={styles.wizardBody}>
            {wizardSubmitted ? (
              /* ===== SUCCESS STATE ===== */
              <div className={styles.wizardSuccess}>
                <div className={styles.wizardSuccessIcon}>
                  <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 48, height: 48 }}>
                    <path d="M22 11.08V12a10 10 0 11-5.93-9.14" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    <path d="M22 4L12 14.01l-3-3" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <h2 style={{ color: '#0f172a', margin: '16px 0 8px' }}>Policy Request Submitted!</h2>
                <p style={{ color: '#64748b', maxWidth: 400, margin: '0 auto 24px', lineHeight: 1.6 }}>
                  Your request for "<strong>{requestForm.policyTitle}</strong>" has been submitted successfully.
                  The Policy Authoring team will be notified and will review your request shortly.
                </p>
                <div style={{ background: '#f0fdfa', borderRadius: 12, padding: 20, maxWidth: 420, margin: '0 auto 24px', textAlign: 'left' as const }}>
                  <div style={{ fontWeight: 600, marginBottom: 12, color: '#0d9488' }}>What happens next?</div>
                  <div style={{ display: 'flex', gap: 12, marginBottom: 10 }}>
                    <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#0d9488', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>1</div>
                    <div style={{ fontSize: 13, color: '#334155' }}>A Policy Author will be assigned to your request</div>
                  </div>
                  <div style={{ display: 'flex', gap: 12, marginBottom: 10 }}>
                    <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#0d9488', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>2</div>
                    <div style={{ fontSize: 13, color: '#334155' }}>They will draft the policy based on your requirements</div>
                  </div>
                  <div style={{ display: 'flex', gap: 12 }}>
                    <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#0d9488', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>3</div>
                    <div style={{ fontSize: 13, color: '#334155' }}>You'll be notified when the draft is ready for review</div>
                  </div>
                </div>
                <button
                  className={styles.wizardBtnPrimary}
                  onClick={() => { setShowRequestWizard(false); resetWizard(); }}
                >
                  Done
                </button>
              </div>
            ) : (
              <>
                {/* ===== STEP 0: Policy Details ===== */}
                {wizardStep === 0 && (
                  <div className={styles.wizardFormGrid}>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel}>Policy Title <span className={styles.wizardRequired}>*</span></label>
                      <input
                        className={styles.wizardInput}
                        type="text"
                        placeholder="e.g. Data Retention Policy for Cloud Storage"
                        value={requestForm.policyTitle}
                        onChange={(e) => updateRequestForm('policyTitle', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel}>Policy Category <span className={styles.wizardRequired}>*</span></label>
                      <select
                        className={styles.wizardSelect}
                        value={requestForm.policyCategory}
                        onChange={(e) => updateRequestForm('policyCategory', e.target.value)}
                      >
                        <option value="">Select category...</option>
                        <option value="IT Security">IT Security</option>
                        <option value="HR Policies">HR Policies</option>
                        <option value="Compliance">Compliance</option>
                        <option value="Health & Safety">Health & Safety</option>
                        <option value="Financial">Financial</option>
                        <option value="Legal">Legal</option>
                        <option value="Environmental">Environmental</option>
                        <option value="Operational">Operational</option>
                        <option value="Data Privacy">Data Privacy</option>
                        <option value="Quality Assurance">Quality Assurance</option>
                      </select>
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel}>Request Type</label>
                      <select
                        className={styles.wizardSelect}
                        value={requestForm.policyType}
                        onChange={(e) => updateRequestForm('policyType', e.target.value)}
                      >
                        <option value="New Policy">New Policy</option>
                        <option value="Policy Update">Policy Update / Revision</option>
                        <option value="Policy Review">Policy Review</option>
                        <option value="Policy Replacement">Policy Replacement</option>
                      </select>
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel}>Priority</label>
                      <select
                        className={styles.wizardSelect}
                        value={requestForm.priority}
                        onChange={(e) => updateRequestForm('priority', e.target.value)}
                      >
                        <option value="Low">Low — No urgency</option>
                        <option value="Medium">Medium — Standard timeline</option>
                        <option value="High">High — Urgent requirement</option>
                        <option value="Critical">Critical — Regulatory deadline</option>
                      </select>
                    </div>
                  </div>
                )}

                {/* ===== STEP 1: Business Case ===== */}
                {wizardStep === 1 && (
                  <div className={styles.wizardFormGrid}>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel}>Business Justification <span className={styles.wizardRequired}>*</span></label>
                      <textarea
                        className={styles.wizardTextarea}
                        rows={5}
                        placeholder="Explain why this policy is needed. Include business drivers, risks of not having this policy, and any relevant context..."
                        value={requestForm.businessJustification}
                        onChange={(e) => updateRequestForm('businessJustification', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel}>Regulatory / Compliance Driver</label>
                      <input
                        className={styles.wizardInput}
                        type="text"
                        placeholder="e.g. GDPR Article 5, ISO 27001, Health & Safety Act"
                        value={requestForm.regulatoryDriver}
                        onChange={(e) => updateRequestForm('regulatoryDriver', e.target.value)}
                      />
                      <div className={styles.wizardHelpText}>If this policy is driven by a regulatory requirement, specify the regulation or standard</div>
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel}>Additional Notes</label>
                      <textarea
                        className={styles.wizardTextarea}
                        rows={3}
                        placeholder="Any additional context, reference documents, or specific requirements for the policy author..."
                        value={requestForm.additionalNotes}
                        onChange={(e) => updateRequestForm('additionalNotes', e.target.value)}
                      />
                    </div>
                  </div>
                )}

                {/* ===== STEP 2: Requirements ===== */}
                {wizardStep === 2 && (
                  <div className={styles.wizardFormGrid}>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel}>Target Audience <span className={styles.wizardRequired}>*</span></label>
                      <input
                        className={styles.wizardInput}
                        type="text"
                        placeholder="e.g. All Employees, IT Department, Management, Contractors"
                        value={requestForm.targetAudience}
                        onChange={(e) => updateRequestForm('targetAudience', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel}>Desired Effective Date</label>
                      <input
                        className={styles.wizardInput}
                        type="date"
                        value={requestForm.desiredEffectiveDate}
                        onChange={(e) => updateRequestForm('desiredEffectiveDate', e.target.value)}
                      />
                    </div>
                    <div className={styles.wizardField}>
                      <label className={styles.wizardLabel}>Read Timeframe (days)</label>
                      <select
                        className={styles.wizardSelect}
                        value={requestForm.readTimeframeDays}
                        onChange={(e) => updateRequestForm('readTimeframeDays', e.target.value)}
                      >
                        <option value="7">7 days</option>
                        <option value="14">14 days</option>
                        <option value="30">30 days</option>
                        <option value="60">60 days</option>
                        <option value="90">90 days</option>
                      </select>
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <div className={styles.wizardCheckboxGroup}>
                        <label className={styles.wizardCheckbox}>
                          <input
                            type="checkbox"
                            checked={requestForm.requiresAcknowledgement}
                            onChange={(e) => updateRequestForm('requiresAcknowledgement', e.target.checked)}
                          />
                          <span>Require employee acknowledgement</span>
                        </label>
                        <label className={styles.wizardCheckbox}>
                          <input
                            type="checkbox"
                            checked={requestForm.requiresQuiz}
                            onChange={(e) => updateRequestForm('requiresQuiz', e.target.checked)}
                          />
                          <span>Require comprehension quiz</span>
                        </label>
                        <label className={styles.wizardCheckbox}>
                          <input
                            type="checkbox"
                            checked={requestForm.notifyAuthors}
                            onChange={(e) => updateRequestForm('notifyAuthors', e.target.checked)}
                          />
                          <span>Notify Policy Authors immediately</span>
                        </label>
                      </div>
                    </div>
                    <div className={styles.wizardFieldFull}>
                      <label className={styles.wizardLabel}>Preferred Author (optional)</label>
                      <input
                        className={styles.wizardInput}
                        type="text"
                        placeholder="Leave blank to auto-assign, or enter a name"
                        value={requestForm.preferredAuthor}
                        onChange={(e) => updateRequestForm('preferredAuthor', e.target.value)}
                      />
                    </div>
                  </div>
                )}

                {/* ===== STEP 3: Review & Submit ===== */}
                {wizardStep === 3 && (
                  <div className={styles.wizardReview}>
                    <div className={styles.wizardReviewSection}>
                      <div className={styles.wizardReviewTitle}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                          <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <path d="M14 2v6h6" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        Policy Details
                      </div>
                      <div className={styles.wizardReviewGrid}>
                        <div><span className={styles.wizardReviewLabel}>Title:</span> <strong>{requestForm.policyTitle}</strong></div>
                        <div><span className={styles.wizardReviewLabel}>Category:</span> {requestForm.policyCategory}</div>
                        <div><span className={styles.wizardReviewLabel}>Type:</span> {requestForm.policyType}</div>
                        <div><span className={styles.wizardReviewLabel}>Priority:</span>
                          <span style={{
                            padding: '2px 10px', borderRadius: 10, fontSize: 12, fontWeight: 600, marginLeft: 6,
                            background: requestForm.priority === 'Critical' ? '#fde7e9' : requestForm.priority === 'High' ? '#fff3e0' : requestForm.priority === 'Medium' ? '#fff8e1' : '#f1f5f9',
                            color: requestForm.priority === 'Critical' ? '#d13438' : requestForm.priority === 'High' ? '#f97316' : requestForm.priority === 'Medium' ? '#f59e0b' : '#64748b'
                          }}>
                            {requestForm.priority}
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className={styles.wizardReviewSection}>
                      <div className={styles.wizardReviewTitle}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                          <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        Business Case
                      </div>
                      <div style={{ fontSize: 13, lineHeight: 1.6, color: '#334155' }}>{requestForm.businessJustification}</div>
                      {requestForm.regulatoryDriver && (
                        <div style={{ marginTop: 8, fontSize: 12, color: '#ef4444' }}>
                          <strong>Regulatory Driver:</strong> {requestForm.regulatoryDriver}
                        </div>
                      )}
                      {requestForm.additionalNotes && (
                        <div style={{ marginTop: 8, fontSize: 12, color: '#64748b', fontStyle: 'italic' }}>
                          <strong>Notes:</strong> {requestForm.additionalNotes}
                        </div>
                      )}
                    </div>

                    <div className={styles.wizardReviewSection}>
                      <div className={styles.wizardReviewTitle}>
                        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                          <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <circle cx="9" cy="7" r="4" stroke="#0d9488" strokeWidth="2"/>
                        </svg>
                        Requirements
                      </div>
                      <div className={styles.wizardReviewGrid}>
                        <div><span className={styles.wizardReviewLabel}>Audience:</span> {requestForm.targetAudience}</div>
                        <div><span className={styles.wizardReviewLabel}>Effective Date:</span> {requestForm.desiredEffectiveDate ? new Date(requestForm.desiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) : 'Not specified'}</div>
                        <div><span className={styles.wizardReviewLabel}>Read Timeframe:</span> {requestForm.readTimeframeDays} days</div>
                        <div><span className={styles.wizardReviewLabel}>Acknowledgement:</span> {requestForm.requiresAcknowledgement ? 'Yes' : 'No'}</div>
                        <div><span className={styles.wizardReviewLabel}>Quiz:</span> {requestForm.requiresQuiz ? 'Yes' : 'No'}</div>
                        {requestForm.preferredAuthor && <div><span className={styles.wizardReviewLabel}>Preferred Author:</span> {requestForm.preferredAuthor}</div>}
                      </div>
                    </div>

                    <div style={{ background: '#fffbeb', borderRadius: 8, padding: 12, border: '1px solid #fde68a', display: 'flex', gap: 10, alignItems: 'flex-start', marginTop: 8 }}>
                      <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 20, height: 20, flexShrink: 0, marginTop: 1 }}>
                        <path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      <div style={{ fontSize: 12, color: '#92400e', lineHeight: 1.5 }}>
                        <strong>Workflow notification:</strong> Upon submission, the Policy Authoring team will receive an email notification with your request details. You will be notified when an author is assigned and when the draft is ready for review.
                      </div>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>

          {/* Wizard Footer */}
          {!wizardSubmitted && (
            <div className={styles.wizardFooter}>
              <div>
                {wizardStep > 0 && (
                  <button className={styles.wizardBtnSecondary} onClick={() => setWizardStep(wizardStep - 1)}>
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                      <path d="M19 12H5M12 19l-7-7 7-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                    Back
                  </button>
                )}
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button className={styles.wizardBtnOutline} onClick={() => setShowRequestWizard(false)}>
                  Cancel
                </button>
                {wizardStep < WIZARD_STEPS.length - 1 ? (
                  <button
                    className={styles.wizardBtnPrimary}
                    onClick={() => setWizardStep(wizardStep + 1)}
                    disabled={
                      (wizardStep === 0 && (!requestForm.policyTitle || !requestForm.policyCategory)) ||
                      (wizardStep === 1 && !requestForm.businessJustification) ||
                      (wizardStep === 2 && !requestForm.targetAudience)
                    }
                  >
                    Next
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                      <path d="M5 12h14M12 5l7 7-7 7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                  </button>
                ) : (
                  <button
                    className={styles.wizardBtnSubmit}
                    onClick={() => setWizardSubmitted(true)}
                  >
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: 14, height: 14 }}>
                      <path d="M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                    Submit Request
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    )}
    </>
  );
};

export default PolicyManagerHeader;
