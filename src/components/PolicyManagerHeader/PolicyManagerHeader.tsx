// @ts-nocheck
import { Icon } from '@fluentui/react/lib/Icon';
import * as React from 'react';
import styles from './PolicyManagerHeader.module.scss';
import { SPFI } from '@pnp/sp';
import { PolicyManagerRole, filterNavForRole, getHeaderVisibility, IRolePermissionEntry } from '../../services/PolicyRoleService';
import { RecentlyViewedService, IRecentlyViewedDisplay } from '../../services/RecentlyViewedService';
import { BookmarkService, IBookmarkEntry } from '../../services/BookmarkService';
import { PolicyRequestWizard } from './PolicyRequestWizard';
import { DwxHubService, DwxNotificationService, DwxNotificationBell } from '@dwx/core';
import { PolicyChatPanel } from '../PolicyChatPanel';
// PolicyHelpPanel deprecated — Help now uses full page at PolicyHelp.aspx

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
  /** DWx Hub service instance for cross-app notifications */
  dwxHub?: DwxHubService;
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
  teamCompliance: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M9 12l2 2 4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  reviews: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2"/>
    </svg>
  ),
  dashboard: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect x="3" y="3" width="7" height="7" rx="1" stroke="currentColor" strokeWidth="2"/>
      <rect x="14" y="3" width="7" height="7" rx="1" stroke="currentColor" strokeWidth="2"/>
      <rect x="3" y="14" width="7" height="7" rx="1" stroke="currentColor" strokeWidth="2"/>
      <rect x="14" y="14" width="7" height="7" rx="1" stroke="currentColor" strokeWidth="2"/>
    </svg>
  ),
  distribution: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M22 2L11 13" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M22 2L15 22l-4-9-9-4 20-7z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  ),
  authorReports: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M14 2v6h6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M8 13h3M8 17h5M8 9h2" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
      <path d="M16 13v4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
    </svg>
  ),
  bulkUpload: (
    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M17 8l-5-5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
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

/**
 * Mapping from header nav item keys to admin toggle keys in PolicyAdmin.
 * Used to apply admin navigation visibility settings from localStorage (pm_nav_visibility).
 * Nav items without a mapping are always shown (e.g. items added dynamically).
 */
const NAV_KEY_TO_TOGGLE_KEY: Record<string, string> = {
  'browse': 'policyHub',
  'my-policies': 'myPolicies',
  'create': 'policyBuilder',
  'author': 'policyAuthor',
  'packs': 'policyPacks',
  'distribution': 'policyDistribution',
  'manager': 'policyManager',
  'analytics': 'policyAnalytics',
  'quiz': 'quizBuilder',
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
  sp,
  dwxHub
}) => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [showProfileDropdown, setShowProfileDropdown] = React.useState(false);
  const [openNavGroup, setOpenNavGroup] = React.useState<string | null>(null);
  const navGroupRef = React.useRef<HTMLDivElement>(null);

  // Close nav dropdown on outside click
  React.useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (navGroupRef.current && !navGroupRef.current.contains(e.target as Node)) {
        setOpenNavGroup(null);
      }
    };
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpenNavGroup(null);
    };
    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
    };
  }, []);
  const [showRecentlyViewedDropdown, setShowRecentlyViewedDropdown] = React.useState(false);

  // Request Policy Wizard — extracted to PolicyRequestWizard.tsx
  const [showRequestWizard, setShowRequestWizard] = React.useState(false);

  // AI Chat Assistant panel
  const [showChatPanel, setShowChatPanel] = React.useState(false);

  // Help Center panel
  // Help panel state removed — Help now navigates to full page

  // Admin navigation visibility toggles (loaded from localStorage, set via PolicyAdmin)
  const [navVisibility, setNavVisibility] = React.useState<Record<string, boolean>>({});

  // Role permissions — loaded from PM_Configuration (set via Admin > Role Permissions)
  const [rolePermissions, setRolePermissions] = React.useState<IRolePermissionEntry[] | null>(null);

  React.useEffect(() => {
    try {
      const saved = localStorage.getItem('pm_nav_visibility');
      if (saved) {
        setNavVisibility(JSON.parse(saved));
      }
    } catch { /* ignore corrupt data */ }

    // Load role permissions from PM_Configuration (with localStorage cache)
    try {
      const cachedPerms = localStorage.getItem('pm_role_permissions');
      if (cachedPerms) {
        setRolePermissions(JSON.parse(cachedPerms));
      }
    } catch { /* ignore */ }

    // Also try loading from SP list for fresh data
    if (sp) {
      sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Admin.RolePermissions.Config'")
        .select('ConfigValue')
        .top(1)()
        .then((items: any[]) => {
          if (items.length > 0 && items[0].ConfigValue) {
            try {
              const perms = JSON.parse(items[0].ConfigValue);
              setRolePermissions(perms);
              localStorage.setItem('pm_role_permissions', items[0].ConfigValue);
            } catch { /* ignore corrupt JSON */ }
          }
        })
        .catch(() => { /* PM_Configuration may not exist — use defaults */ });
    }
  }, []);

  // Cross-app notification service (from DWx Hub)
  const dwxNotificationService = React.useMemo(() => {
    if (dwxHub) {
      try { return new DwxNotificationService(dwxHub); } catch { return null; }
    }
    return null;
  }, [dwxHub]);

  // Refs for click-outside detection
  const profileRef = React.useRef<HTMLDivElement>(null);
  const recentlyViewedRef = React.useRef<HTMLDivElement>(null);
  const bookmarksRef = React.useRef<HTMLDivElement>(null);

  // Recently viewed policies — loaded from localStorage via RecentlyViewedService
  const [recentlyViewedPolicies, setRecentlyViewedPolicies] = React.useState<IRecentlyViewedDisplay[]>([]);

  // Bookmarked policies — loaded from localStorage via BookmarkService
  const [showBookmarksDropdown, setShowBookmarksDropdown] = React.useState(false);
  const [bookmarkedPolicies, setBookmarkedPolicies] = React.useState<IBookmarkEntry[]>([]);

  // Local notification bell — reads from PM_Notifications
  const [showNotifDropdown, setShowNotifDropdown] = React.useState(false);
  const [localNotifications, setLocalNotifications] = React.useState<Array<{ Id: number; Title: string; Message: string; Type: string; ActionUrl: string; Created: string }>>([]);
  const [localUnreadCount, setLocalUnreadCount] = React.useState(0);
  const notifRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    setRecentlyViewedPolicies(RecentlyViewedService.getRecentlyViewed(5));
  }, []);

  // Refresh the list every time the dropdown is opened (picks up views from other pages)
  React.useEffect(() => {
    if (showRecentlyViewedDropdown) {
      setRecentlyViewedPolicies(RecentlyViewedService.getRecentlyViewed(5));
    }
  }, [showRecentlyViewedDropdown]);

  React.useEffect(() => {
    if (showBookmarksDropdown) {
      setBookmarkedPolicies(BookmarkService.getAll());
    }
  }, [showBookmarksDropdown]);

  // Load local notifications from PM_Notifications on mount and when dropdown opens
  React.useEffect(() => {
    if (!sp || !userEmail) return;
    const loadNotifications = async () => {
      try {
        // Get unread count
        const unreadItems: any[] = await sp.web.lists.getByTitle('PM_Notifications')
          .items.filter(`RecipientEmail eq '${userEmail.replace(/'/g, "''")}' and IsRead eq 0`)
          .select('Id')
          .top(50)();
        setLocalUnreadCount(unreadItems.length);

        // Get 10 most recent for the dropdown
        if (showNotifDropdown) {
          const recent: any[] = await sp.web.lists.getByTitle('PM_Notifications')
            .items.filter(`RecipientEmail eq '${userEmail.replace(/'/g, "''")}'`)
            .select('Id', 'Title', 'Message', 'Type', 'ActionUrl', 'Created', 'IsRead')
            .orderBy('Created', false)
            .top(10)();
          setLocalNotifications(recent);
        }
      } catch {
        // PM_Notifications may not exist yet
      }
    };
    loadNotifications();
  }, [sp, userEmail, showNotifDropdown]);

  // Click-outside handler to close dropdowns
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (profileRef.current && !profileRef.current.contains(event.target as Node)) {
        setShowProfileDropdown(false);
      }
      if (recentlyViewedRef.current && !recentlyViewedRef.current.contains(event.target as Node)) {
        setShowRecentlyViewedDropdown(false);
      }
      if (bookmarksRef.current && !bookmarksRef.current.contains(event.target as Node)) {
        setShowBookmarksDropdown(false);
      }
      if (notifRef.current && !notifRef.current.contains(event.target as Node)) {
        setShowNotifDropdown(false);
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

  const [headerSearchResults, setHeaderSearchResults] = React.useState<any[]>([]);
  const [headerSearching, setHeaderSearching] = React.useState(false);
  const [showSearchDropdown, setShowSearchDropdown] = React.useState(false);
  const searchDebounceRef = React.useRef<any>(null);
  const searchDropdownRef = React.useRef<HTMLDivElement>(null);

  // Close search dropdown on outside click
  React.useEffect(() => {
    const handleClickOutsideSearch = (e: MouseEvent) => {
      if (searchDropdownRef.current && !searchDropdownRef.current.contains(e.target as Node)) {
        setShowSearchDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutsideSearch);
    return () => document.removeEventListener('mousedown', handleClickOutsideSearch);
  }, []);

  // Debounced inline search — queries PM_Policies + secure document libraries
  const handleSearchInputChange = (value: string) => {
    setSearchQuery(value);
    if (searchDebounceRef.current) clearTimeout(searchDebounceRef.current);
    if (!value.trim() || value.trim().length < 2) {
      setHeaderSearchResults([]);
      setShowSearchDropdown(false);
      return;
    }
    searchDebounceRef.current = setTimeout(async () => {
      if (!sp) return;
      setHeaderSearching(true);
      setShowSearchDropdown(true);
      try {
        // Primary query: PM_Policies list (client-side filter for reliability)
        const allItems = await sp.web.lists.getByTitle('PM_Policies')
          .items.select('Id', 'Title', 'PolicyName', 'PolicyNumber', 'PolicyCategory', 'PolicyStatus', 'PolicyDescription')
          .orderBy('Title')
          .top(500)();

        const q = value.trim().toLowerCase();
        const filtered = allItems.filter((p: any) =>
          (p.PolicyName || p.Title || '').toLowerCase().includes(q) ||
          (p.PolicyNumber || '').toLowerCase().includes(q) ||
          (p.PolicyCategory || '').toLowerCase().includes(q) ||
          (p.PolicyDescription || '').toLowerCase().includes(q)
        ).slice(0, 8);

        // Also search secure document libraries if configured
        let secureResults: any[] = [];
        try {
          const configItems = await sp.web.lists.getByTitle('PM_Configuration')
            .items.filter("ConfigKey eq 'Admin.SecureLibraries.Config'")
            .select('ConfigValue')
            .top(1)();
          if (configItems.length > 0 && configItems[0].ConfigValue) {
            const secureLibs = JSON.parse(configItems[0].ConfigValue);
            for (const lib of secureLibs.filter((l: any) => l.isActive)) {
              try {
                const libItems = await sp.web.lists.getByTitle(lib.title || lib.libraryName)
                  .items.select('Id', 'Title', 'FileLeafRef')
                  .top(50)();
                const libMatches = libItems.filter((item: any) =>
                  (item.Title || item.FileLeafRef || '').toLowerCase().includes(q)
                ).slice(0, 3).map((item: any) => ({
                  ...item,
                  PolicyName: item.Title || item.FileLeafRef,
                  PolicyNumber: '',
                  PolicyCategory: lib.title,
                  PolicyStatus: 'Secure',
                  _isSecureLib: true,
                  _libraryUrl: lib.libraryUrl
                }));
                secureResults = [...secureResults, ...libMatches];
              } catch { /* library may not exist or no access */ }
            }
          }
        } catch { /* PM_Configuration may not exist */ }

        setHeaderSearchResults([...filtered, ...secureResults].slice(0, 10));
      } catch {
        setHeaderSearchResults([]);
      }
      setHeaderSearching(false);
    }, 300);
  };

  const handleSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      if (onSearch) {
        onSearch(searchQuery);
      }
      setShowSearchDropdown(false);
      // Navigate to search page with query parameter
      if (searchQuery.trim()) {
        window.location.href = `/sites/PolicyManager/SitePages/PolicySearch.aspx?q=${encodeURIComponent(searchQuery.trim())}`;
      }
    }
    if (e.key === 'Escape') {
      setShowSearchDropdown(false);
    }
  };

  const siteUrl = sp ? '' : '';

  // Format login time
  const displayLoginTime = loginTime || new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  // ── Grouped nav structure ──
  // Flat items: visible to all (role-filtered). Dropdown groups: Author, Manager.
  interface INavGroup {
    key: string;
    text: string;
    icon: JSX.Element;
    minRole: 'User' | 'Author' | 'Manager' | 'Admin';
    children: INavItem[];
  }

  const flatNavItems: INavItem[] = [
    { key: 'my-policies', text: 'My Policies', icon: NavIcons.authored, href: '/sites/PolicyManager/SitePages/MyPolicies.aspx' },
    { key: 'browse', text: 'Policy Hub', icon: NavIcons.browse, href: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
  ];

  const navGroups: INavGroup[] = [
    {
      key: 'author-group', text: 'Author', icon: NavIcons.authored, minRole: 'Author',
      children: [
        { key: 'newpolicy', text: 'New Policy', icon: NavIcons.create, href: '/sites/PolicyManager/SitePages/PolicyBuilder.aspx' },
        { key: 'author', text: 'Drafts & Pipeline', icon: NavIcons.authored, href: '/sites/PolicyManager/SitePages/PolicyAuthor.aspx' },
        { key: 'packs', text: 'Policy Packs', icon: NavIcons.packs, href: '/sites/PolicyManager/SitePages/PolicyPacks.aspx' },
        { key: 'quiz', text: 'Quiz Builder', icon: NavIcons.quiz, href: '/sites/PolicyManager/SitePages/QuizBuilder.aspx' },
        { key: 'author-reports', text: 'Reports', icon: NavIcons.authorReports, href: '/sites/PolicyManager/SitePages/PolicyAuthorReports.aspx' },
        { key: 'bulk-upload', text: 'Bulk Upload', icon: NavIcons.bulkUpload, href: '/sites/PolicyManager/SitePages/PolicyBulkUpload.aspx' },
      ]
    },
    {
      key: 'manager-group', text: 'Manager', icon: NavIcons.manager, minRole: 'Manager',
      children: [
        { key: 'manager-dashboard', text: 'Dashboard', icon: NavIcons.dashboard, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx' },
        { key: 'approvals', text: 'Approvals', icon: NavIcons.approvals, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=approvals' },
        { key: 'distribution', text: 'Distribution', icon: NavIcons.distribution, href: '/sites/PolicyManager/SitePages/PolicyDistribution.aspx' },
        { key: 'team-compliance', text: 'Team Compliance', icon: NavIcons.teamCompliance, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=team-compliance' },
        { key: 'delegations', text: 'Delegations', icon: NavIcons.delegations, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=delegations' },
        { key: 'reviews', text: 'Review Cycles', icon: NavIcons.reviews, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=reviews' },
        { key: 'reports', text: 'Reports', icon: NavIcons.analytics, href: '/sites/PolicyManager/SitePages/PolicyManagerView.aspx?tab=reports' },
        { key: 'analytics', text: 'Analytics', icon: NavIcons.analytics, href: '/sites/PolicyManager/SitePages/PolicyAnalytics.aspx' },
        { key: 'request-policy', text: 'Request Policy', icon: NavIcons.create, href: '#request-policy' },
      ]
    }
  ];

  // ── Secure Policies dropdown — group-membership based ──
  const [secureLibItems, setSecureLibItems] = React.useState<Array<{ title: string; libraryUrl: string; icon: string }>>([]);
  const [secureLibsChecked, setSecureLibsChecked] = React.useState(false);

  React.useEffect(() => {
    if (secureLibsChecked || !sp) return;
    setSecureLibsChecked(true);

    // Load secure libraries config from localStorage (fast) then SP (authoritative)
    let libs: Array<{ title: string; libraryUrl: string; securityGroups: string[]; icon: string; isActive: boolean }> = [];
    try {
      const cached = localStorage.getItem('pm_secure_libraries');
      if (cached) libs = JSON.parse(cached).filter((l: any) => l.isActive);
    } catch { /* */ }

    // Also try loading from SP for fresh data
    sp.web.lists.getByTitle('PM_Configuration')
      .items.filter("ConfigKey eq 'Admin.SecureLibraries.Config'")
      .select('ConfigValue').top(1)()
      .then((items: any[]) => {
        if (items.length > 0 && items[0].ConfigValue) {
          try {
            libs = JSON.parse(items[0].ConfigValue).filter((l: any) => l.isActive);
            localStorage.setItem('pm_secure_libraries', items[0].ConfigValue);
          } catch { /* */ }
        }
        if (libs.length === 0) return;

        // Check which groups the current user belongs to
        sp.web.currentUser.groups()
          .then((userGroups: any[]) => {
            const userGroupNames = new Set(userGroups.map((g: any) => g.Title));
            const accessible = libs.filter(lib =>
              lib.securityGroups.some(sg => userGroupNames.has(sg))
            );
            if (accessible.length > 0) {
              setSecureLibItems(accessible.map(l => ({ title: l.title, libraryUrl: l.libraryUrl, icon: l.icon || 'Lock' })));
            }
          })
          .catch(() => { /* can't check groups */ });
      })
      .catch(() => { /* PM_Configuration may not exist */ });
  }, [sp, secureLibsChecked]);

  // Legacy flat list for role filtering compatibility
  const allFlatKeys = [
    ...flatNavItems,
    ...navGroups.flatMap(g => g.children)
  ];
  const allNavItems = navItems.length > 0 ? navItems : allFlatKeys;
  const roleFiltered = policyRole ? filterNavForRole(allNavItems, policyRole, rolePermissions) : allNavItems;
  const roleFilteredKeys = new Set(roleFiltered.map(item => item.key));

  // Apply admin navigation toggles
  const isNavVisible = (key: string): boolean => {
    if (!roleFilteredKeys.has(key)) return false;
    if (Object.keys(navVisibility).length === 0) return true;
    const toggleKey = NAV_KEY_TO_TOGGLE_KEY[key];
    if (!toggleKey) return true;
    return navVisibility[toggleKey] !== false;
  };

  const displayFlatItems = flatNavItems.filter(item => isNavVisible(item.key));
  const displayGroups = navGroups
    .filter(g => {
      if (!policyRole) return true;
      const roleLevel: Record<string, number> = { User: 0, Author: 1, Manager: 2, Admin: 3 };
      return (roleLevel[policyRole] || 0) >= (roleLevel[g.minRole] || 0);
    })
    .map(g => ({ ...g, children: g.children.filter(c => isNavVisible(c.key)) }))
    .filter(g => g.children.length > 0);

  // Override header visibility based on role
  const roleVisibility = policyRole ? getHeaderVisibility(policyRole, rolePermissions) : null;

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

  // Handle help icon click — navigate to full Help Center page
  const handleHelpClick = () => {
    if (onHelpClick) {
      onHelpClick();
    } else {
      window.location.href = '/sites/PolicyManager/SitePages/PolicyHelp.aspx';
    }
  };

  // Handle profile avatar click
  const handleProfileClick = () => {
    setShowProfileDropdown(!showProfileDropdown);
  };

  return (
    <>
    <header className={styles.headerContainer}>
      {/* Top Bar - Dark teal gradient */}
      <div className={styles.topBar}>
        {/* Logo Section */}
        <a href="/sites/PolicyManager/SitePages/PolicyHub.aspx" className={styles.logoSection} style={{ textDecoration: 'none', color: 'inherit' }}>
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
        </a>

        {/* Search Section */}
        {showSearch && (
          <div className={styles.searchSection} ref={searchDropdownRef} style={{ position: 'relative' }}>
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
                onChange={(e) => handleSearchInputChange(e.target.value)}
                onKeyDown={handleSearchKeyDown}
                onFocus={() => { if (headerSearchResults.length > 0) setShowSearchDropdown(true); }}
              />
            </div>
            {/* Inline Search Results Dropdown */}
            {showSearchDropdown && (
              <div style={{
                position: 'absolute', top: '100%', left: 0, right: 0, marginTop: 4,
                background: '#fff', borderRadius: 6, boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
                border: '1px solid #e2e8f0', zIndex: 1000, maxHeight: 400, overflowY: 'auto'
              }}>
                {headerSearching ? (
                  <div style={{ padding: '16px 20px', textAlign: 'center' }}>
                    <span style={{ fontSize: 12, color: '#94a3b8' }}>Searching...</span>
                  </div>
                ) : headerSearchResults.length === 0 ? (
                  <div style={{ padding: '16px 20px', textAlign: 'center' }}>
                    <span style={{ fontSize: 12, color: '#94a3b8' }}>No policies found for &ldquo;{searchQuery}&rdquo;</span>
                  </div>
                ) : (
                  <>
                    <div style={{ padding: '8px 16px', borderBottom: '1px solid #f1f5f9' }}>
                      <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 500 }}>{headerSearchResults.length} result{headerSearchResults.length !== 1 ? 's' : ''}</span>
                    </div>
                    {headerSearchResults.map((policy: any) => {
                      const statusColors: Record<string, { bg: string; color: string }> = {
                        Published: { bg: '#dcfce7', color: '#16a34a' },
                        Draft: { bg: '#f1f5f9', color: '#64748b' },
                        'In Review': { bg: '#fef3c7', color: '#d97706' },
                        Archived: { bg: '#f1f5f9', color: '#94a3b8' },
                        Secure: { bg: '#ede9fe', color: '#7c3aed' }
                      };
                      const sc = statusColors[policy.PolicyStatus] || statusColors.Draft;
                      return (
                        <div
                          key={policy.Id}
                          role="button"
                          tabIndex={0}
                          onClick={() => {
                            setShowSearchDropdown(false);
                            setSearchQuery('');
                            window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&highlight=true`;
                          }}
                          onKeyDown={(ev) => { if (ev.key === 'Enter') { setShowSearchDropdown(false); window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&highlight=true`; } }}
                          style={{
                            padding: '10px 16px', cursor: 'pointer', borderBottom: '1px solid #f8fafc',
                            display: 'flex', alignItems: 'center', gap: 10,
                            transition: 'background 0.1s'
                          }}
                          onMouseEnter={(ev) => { (ev.currentTarget as HTMLElement).style.background = '#f0fdfa'; }}
                          onMouseLeave={(ev) => { (ev.currentTarget as HTMLElement).style.background = '#fff'; }}
                        >
                          <div style={{ width: 32, height: 32, borderRadius: 4, background: '#f0fdfa', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                            <svg viewBox="0 0 24 24" fill="none" style={{ width: 16, height: 16 }}>
                              <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                            </svg>
                          </div>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {policy.PolicyName || policy.Title}
                            </div>
                            <div style={{ fontSize: 11, color: '#94a3b8' }}>
                              {policy.PolicyNumber} | {policy.PolicyCategory}
                            </div>
                          </div>
                          <span style={{
                            fontSize: 9, fontWeight: 600, padding: '2px 6px', borderRadius: 3,
                            background: sc.bg, color: sc.color, flexShrink: 0
                          }}>
                            {policy.PolicyStatus}
                          </span>
                        </div>
                      );
                    })}
                    <div
                      role="button" tabIndex={0}
                      onClick={() => {
                        setShowSearchDropdown(false);
                        window.location.href = `/sites/PolicyManager/SitePages/PolicySearch.aspx?q=${encodeURIComponent(searchQuery.trim())}`;
                      }}
                      onKeyDown={(ev) => { if (ev.key === 'Enter') { setShowSearchDropdown(false); window.location.href = `/sites/PolicyManager/SitePages/PolicySearch.aspx?q=${encodeURIComponent(searchQuery.trim())}`; } }}
                      style={{ padding: '10px 16px', textAlign: 'center', cursor: 'pointer', borderTop: '1px solid #e2e8f0', background: '#f8fafc' }}
                      onMouseEnter={(ev) => { (ev.currentTarget as HTMLElement).style.background = '#f0fdfa'; }}
                      onMouseLeave={(ev) => { (ev.currentTarget as HTMLElement).style.background = '#f8fafc'; }}
                    >
                      <span style={{ fontSize: 12, color: '#0d9488', fontWeight: 600 }}>
                        View all results for &ldquo;{searchQuery}&rdquo; →
                      </span>
                    </div>
                  </>
                )}
              </div>
            )}
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
            {/* Request Policy moved to Manager dropdown */}
            <div className={styles.dropdownContainer} ref={recentlyViewedRef} style={{ display: 'inline-flex' }}>
              <button
                className={styles.quickActionBtn}
                type="button"
                title="Recently Viewed"
                aria-label="Recently Viewed"
                onClick={() => {
                  setShowRecentlyViewedDropdown(!showRecentlyViewedDropdown);
                  setShowProfileDropdown(false);
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
                    {recentlyViewedPolicies.length === 0 ? (
                      <div style={{ padding: '16px 20px', textAlign: 'center', color: '#605e5c', fontSize: '13px' }}>
                        No recently viewed policies yet. Browse or view a policy to see it here.
                      </div>
                    ) : (
                      recentlyViewedPolicies.map(policy => (
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
                      ))
                    )}
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

          {/* Bookmarks dropdown */}
          <div className={styles.dropdownContainer} ref={bookmarksRef} style={{ display: 'inline-flex' }}>
            <button
              className={styles.actionButton}
              type="button"
              title="Bookmarks"
              aria-label="Bookmarked Policies"
              onClick={() => {
                setShowBookmarksDropdown(!showBookmarksDropdown);
                setShowRecentlyViewedDropdown(false);
                setShowProfileDropdown(false);
              }}
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M19 21l-7-5-7 5V5a2 2 0 012-2h10a2 2 0 012 2z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </button>

            {showBookmarksDropdown && (
              <div className={styles.dropdownPanel} style={{ left: 0, right: 'auto', minWidth: '340px' }}>
                <div className={styles.dropdownArrow} style={{ left: '40px', right: 'auto' }} />
                <div className={styles.dropdownPanelHeader}>
                  <span className={styles.dropdownPanelTitle}>Bookmarked Policies</span>
                  <span className={styles.dropdownPanelBadge}>{bookmarkedPolicies.length} saved</span>
                </div>
                <div className={styles.dropdownPanelBody}>
                  {bookmarkedPolicies.length === 0 ? (
                    <div style={{ padding: '16px 20px', textAlign: 'center', color: '#605e5c', fontSize: '13px' }}>
                      No bookmarked policies yet. Click the bookmark icon on any policy to save it here.
                    </div>
                  ) : (
                    bookmarkedPolicies.map(policy => (
                      <a
                        key={policy.id}
                        href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.id}&mode=browse`}
                        className={styles.notificationItem}
                        style={{ textDecoration: 'none', color: 'inherit' }}
                        onClick={() => setShowBookmarksDropdown(false)}
                      >
                        <div
                          className={styles.notificationItemIcon}
                          style={{ background: '#0d9488' }}
                        >
                          <svg viewBox="0 0 24 24" fill="white" xmlns="http://www.w3.org/2000/svg" style={{ width: 16, height: 16 }}>
                            <path d="M19 21l-7-5-7 5V5a2 2 0 012-2h10a2 2 0 012 2z" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          </svg>
                        </div>
                        <div className={styles.notificationItemContent}>
                          <span className={styles.notificationItemTitle}>{policy.title}</span>
                          <span className={styles.notificationItemMessage}>{policy.category || 'Policy'}</span>
                        </div>
                        <button
                          type="button"
                          title="Remove bookmark"
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', padding: 4, flexShrink: 0 }}
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            BookmarkService.remove(policy.id);
                            setBookmarkedPolicies(BookmarkService.getAll());
                          }}
                        >
                          <svg viewBox="0 0 24 24" fill="none" width="14" height="14"><path d="M6 18L18 6M6 6l12 12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                        </button>
                      </a>
                    ))
                  )}
                </div>
                {bookmarkedPolicies.length > 0 && (
                  <div className={styles.dropdownPanelFooter}>
                    <button
                      type="button"
                      onClick={() => { BookmarkService.clearAll(); setBookmarkedPolicies([]); }}
                      className={styles.dropdownFooterLink}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#0d9488', fontSize: 12 }}
                    >
                      Clear All Bookmarks
                    </button>
                  </div>
                )}
              </div>
            )}
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

          {/* AI Chat Assistant */}
          <button
            className={styles.actionButton}
            type="button"
            title="AI Assistant"
            onClick={() => setShowChatPanel(true)}
            aria-label="AI Assistant"
          >
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
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

          {/* Admin Settings (Cog) — visible only if role allows it (role takes precedence over prop) */}
          {(roleVisibility ? roleVisibility.showSettings : showSettings) && (
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
              title="Admin Centre"
              aria-label="Admin Centre"
            >
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M12 15a3 3 0 100-6 3 3 0 000 6z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z" stroke="currentColor" strokeWidth="2"/>
              </svg>
            </button>
          )}

          {/* Local Notification Bell — reads from PM_Notifications */}
          {showNotifications && sp && (
            <div className={styles.dropdownContainer} ref={notifRef} style={{ position: 'relative' }}>
              <button
                type="button"
                className={styles.actionButton}
                onClick={() => setShowNotifDropdown(!showNotifDropdown)}
                title="Notifications"
                aria-label={`Notifications${localUnreadCount > 0 ? ` (${localUnreadCount} unread)` : ''}`}
                style={{ position: 'relative' }}
              >
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M18 8A6 6 0 106 8c0 7-3 9-3 9h18s-3-2-3-9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  <path d="M13.73 21a2 2 0 01-3.46 0" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
                {localUnreadCount > 0 && (
                  <span style={{
                    position: 'absolute', top: 0, right: 0,
                    background: '#dc2626', color: '#fff', fontSize: 9, fontWeight: 700,
                    minWidth: 16, height: 16, borderRadius: 8,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    padding: '0 4px', lineHeight: 1, border: '2px solid #0f766e',
                  }}>
                    {localUnreadCount > 9 ? '9+' : localUnreadCount}
                  </span>
                )}
              </button>
              {showNotifDropdown && (
                <div className={styles.dropdownPanel} style={{ width: 340, right: 0, maxHeight: 420, overflowY: 'auto' }}>
                  <div className={styles.dropdownArrow} />
                  <div style={{ padding: '12px 16px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: '#0f172a' }}>Notifications</span>
                    {localUnreadCount > 0 && (
                      <span style={{ fontSize: 10, color: '#0d9488', fontWeight: 600, cursor: 'pointer' }}
                        onClick={async () => {
                          try {
                            for (const n of localNotifications.filter((x: any) => !x.IsRead)) {
                              await sp.web.lists.getByTitle('PM_Notifications').items.getById(n.Id).update({ IsRead: true });
                            }
                            setLocalUnreadCount(0);
                            setLocalNotifications(localNotifications.map((n: any) => ({ ...n, IsRead: true })));
                          } catch { /* best effort */ }
                        }}
                        role="button" tabIndex={0}
                      >Mark all read</span>
                    )}
                  </div>
                  {localNotifications.length === 0 ? (
                    <div style={{ padding: '24px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 12 }}>No notifications</div>
                  ) : (
                    localNotifications.map((n: any) => {
                      const notifColor = getNotificationColor(n.Type || 'task');
                      const timeAgo = n.Created ? (() => {
                        const mins = Math.round((Date.now() - new Date(n.Created).getTime()) / 60000);
                        if (mins < 60) return `${mins}m ago`;
                        if (mins < 1440) return `${Math.round(mins / 60)}h ago`;
                        return `${Math.round(mins / 1440)}d ago`;
                      })() : '';
                      return (
                        <a
                          key={n.Id}
                          href={n.ActionUrl || '#'}
                          style={{
                            display: 'flex', gap: 10, padding: '10px 16px', textDecoration: 'none',
                            borderBottom: '1px solid #f8fafc', cursor: 'pointer',
                            background: n.IsRead ? 'transparent' : '#f0fdfa',
                          }}
                          onClick={() => {
                            if (!n.IsRead) {
                              try { sp.web.lists.getByTitle('PM_Notifications').items.getById(n.Id).update({ IsRead: true }); } catch { /* */ }
                              setLocalUnreadCount(Math.max(0, localUnreadCount - 1));
                            }
                          }}
                        >
                          <div style={{
                            width: 28, height: 28, borderRadius: '50%', flexShrink: 0,
                            background: `${notifColor}15`, color: notifColor,
                            display: 'flex', alignItems: 'center', justifyContent: 'center', marginTop: 2,
                          }}>
                            <svg viewBox="0 0 24 24" fill="none" style={{ width: 14, height: 14 }}>
                              <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2"/>
                              <path d="M12 8v4l3 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                            </svg>
                          </div>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontSize: 12, fontWeight: n.IsRead ? 400 : 600, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{n.Title}</div>
                            <div style={{ fontSize: 11, color: '#64748b', marginTop: 1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{n.Message}</div>
                            <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>{timeAgo}</div>
                          </div>
                          {!n.IsRead && (
                            <div style={{ width: 6, height: 6, borderRadius: '50%', background: '#0d9488', flexShrink: 0, marginTop: 8 }} />
                          )}
                        </a>
                      );
                    })
                  )}
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
                  {policyRole && (
                    <span style={{
                      display: 'inline-block', marginTop: 4,
                      padding: '2px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                      background: policyRole === 'Admin' ? '#fef2f2' : policyRole === 'Manager' ? '#fffbeb' : policyRole === 'Author' ? '#f0fdf4' : '#f0f9ff',
                      color: policyRole === 'Admin' ? '#dc2626' : policyRole === 'Manager' ? '#d97706' : policyRole === 'Author' ? '#16a34a' : '#0284c7'
                    }}>
                      {policyRole}
                    </span>
                  )}
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
        <div className={styles.navItems} ref={navGroupRef}>
          {/* Flat nav items */}
          {displayFlatItems.map((item) => (
            <a
              key={item.key}
              href={item.href}
              className={`${styles.navItem} ${activeNavKey === item.key ? styles.active : ''}`}
              onClick={item.onClick}
            >
              {item.icon}
              {item.text}
            </a>
          ))}

          {/* Separator between flat and grouped items */}
          {displayGroups.length > 0 && (
            <div style={{ width: 1, height: 20, background: '#e2e8f0', margin: '8px 4px', alignSelf: 'center' }} />
          )}

          {/* Dropdown groups */}
          {displayGroups.map((group) => {
            const isOpen = openNavGroup === group.key;
            const hasActiveChild = group.children.some(c => activeNavKey === c.key);
            return (
              <div key={group.key} style={{ position: 'relative' }}>
                <button
                  type="button"
                  className={`${styles.navItem} ${hasActiveChild ? styles.active : ''}`}
                  onClick={() => setOpenNavGroup(isOpen ? null : group.key)}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'inherit', fontSize: 'inherit' }}
                >
                  {group.icon}
                  {group.text}
                  <svg viewBox="0 0 24 24" fill="none" style={{
                    width: 12, height: 12, marginLeft: 2, transition: 'transform 0.2s',
                    transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)'
                  }}>
                    <path d="M6 9l6 6 6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </button>

                {isOpen && (
                  <div style={{
                    position: 'absolute', top: '100%', left: 0, background: '#fff',
                    border: '1px solid #e2e8f0', borderRadius: 4,
                    boxShadow: '0 8px 24px rgba(0,0,0,0.12)', minWidth: 240,
                    zIndex: 100, padding: '4px 0'
                  }}>
                    {group.children.map((child) => (
                      <a
                        key={child.key}
                        href={child.key === 'request-policy' ? undefined : child.href}
                        onClick={(e) => {
                          setOpenNavGroup(null);
                          if (child.key === 'request-policy') {
                            e.preventDefault();
                            setShowRequestWizard(true);
                          }
                        }}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 10, padding: '9px 16px',
                          fontSize: 13, color: activeNavKey === child.key ? '#0d9488' : '#334155',
                          fontWeight: activeNavKey === child.key ? 600 : 400,
                          background: activeNavKey === child.key ? '#f0fdfa' : 'transparent',
                          textDecoration: 'none', cursor: 'pointer'
                        }}
                        onMouseEnter={(e) => { if (activeNavKey !== child.key) (e.currentTarget as HTMLElement).style.background = '#f0fdfa'; }}
                        onMouseLeave={(e) => { if (activeNavKey !== child.key) (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
                      >
                        <span style={{ width: 20, display: 'flex', justifyContent: 'center' }}>{child.icon}</span>
                        {child.text}
                      </a>
                    ))}
                    {/* Request Policy now included in manager-group children */}
                  </div>
                )}
              </div>
            );
          })}

          {/* Secure Policies dropdown — only visible to security group members */}
          {secureLibItems.length > 0 && (
            <>
              <div style={{ width: 1, height: 20, background: '#e2e8f0', margin: '8px 4px', alignSelf: 'center' }} />
              <div style={{ position: 'relative' }}>
                <button
                  type="button"
                  className={`${styles.navItem} ${activeNavKey === 'secure' ? styles.active : ''}`}
                  onClick={() => setOpenNavGroup(openNavGroup === 'secure-group' ? null : 'secure-group')}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', fontFamily: 'inherit', fontSize: 'inherit' }}
                >
                  <svg viewBox="0 0 24 24" fill="none" style={{ width: 16, height: 16 }}>
                    <rect x="3" y="11" width="18" height="11" rx="2" stroke="currentColor" strokeWidth="2"/>
                    <path d="M7 11V7a5 5 0 0110 0v4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                  </svg>
                  Secure Policies
                  <svg viewBox="0 0 24 24" fill="none" style={{
                    width: 12, height: 12, marginLeft: 2, transition: 'transform 0.2s',
                    transform: openNavGroup === 'secure-group' ? 'rotate(180deg)' : 'rotate(0deg)'
                  }}>
                    <path d="M6 9l6 6 6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </button>

                {openNavGroup === 'secure-group' && (
                  <div style={{
                    position: 'absolute', top: '100%', left: 0, background: '#fff',
                    border: '1px solid #e2e8f0', borderRadius: 4,
                    boxShadow: '0 8px 24px rgba(0,0,0,0.12)', minWidth: 260,
                    zIndex: 100, padding: '4px 0'
                  }}>
                    {secureLibItems.map(lib => (
                      <a
                        key={lib.libraryUrl}
                        href={`/sites/PolicyManager/SitePages/PolicyHub.aspx?library=${encodeURIComponent(lib.libraryUrl)}&title=${encodeURIComponent(lib.title)}`}
                        onClick={() => setOpenNavGroup(null)}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 10, padding: '9px 16px',
                          fontSize: 13, color: '#334155', textDecoration: 'none', cursor: 'pointer'
                        }}
                        onMouseEnter={(e) => (e.currentTarget as HTMLElement).style.background = '#f0fdfa'}
                        onMouseLeave={(e) => (e.currentTarget as HTMLElement).style.background = 'transparent'}
                      >
                        <Icon iconName={lib.icon || 'Lock'} styles={{ root: { fontSize: 16, color: '#0d9488' } }} />
                        {lib.title}
                      </a>
                    ))}
                  </div>
                )}
              </div>
            </>
          )}
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

      {/* Breadcrumbs Bar — hidden per user request; keep code for future re-enable */}

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

    {/* AI Chat Assistant Panel */}
    {sp && (
      <PolicyChatPanel
        isOpen={showChatPanel}
        onDismiss={() => setShowChatPanel(false)}
        sp={sp}
        userRole={policyRole || PolicyManagerRole.User}
        userName={userName || ''}
      />
    )}

    {/* Help Center — now uses full page at PolicyHelp.aspx (panel deprecated) */}

    </>
  );
};

export default PolicyManagerHeader;
