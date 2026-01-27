// User Preferences Models
// Interfaces for personalization and user settings

import { IBaseListItem } from './ICommon';

/**
 * User Preferences
 */
export interface IUserPreferences extends IBaseListItem {
  UserId: string;
  UserEmail: string;
  DisplayName: string;

  // Dashboard Layout
  DashboardLayout?: IDashboardLayout;

  // Favorites
  FavoriteProcesses?: number[]; // Array of process IDs
  FavoriteTemplates?: number[]; // Array of template IDs
  FavoriteViews?: string[]; // Array of saved view names

  // Theme & Branding
  ThemePreference?: IThemePreference;
  CustomTheme?: ICustomTheme; // Deprecated - kept for backward compatibility
  CustomThemes?: ICustomTheme[]; // Array of saved themes
  ActiveThemeId?: string; // ID of currently active theme

  // Saved Filters & Views
  SavedFilters?: ISavedFilter[];
  SavedViews?: ISavedView[];
  DefaultView?: string; // Default view name

  // Notification Preferences
  NotificationSettings?: INotificationSettings;

  // Localization
  Language?: string; // 'en-US', 'es-ES', 'fr-FR', etc.
  TimeZone?: string; // IANA timezone (e.g., 'America/New_York')
  DateFormat?: string; // 'MM/DD/YYYY', 'DD/MM/YYYY', 'YYYY-MM-DD'
  TimeFormat?: '12h' | '24h';
  NumberFormat?: string; // 'en-US', 'de-DE', etc.

  // Display Preferences
  DensityMode?: 'compact' | 'normal' | 'comfortable';
  DefaultItemsPerPage?: number;
  ShowWelcomeMessage?: boolean;
  ShowTipsAndTricks?: boolean;

  // Advanced Preferences
  KeyboardShortcutsEnabled?: boolean;
  AnimationsEnabled?: boolean;
  AccessibilityMode?: boolean;
  HighContrastMode?: boolean;

  // Data
  Preferences?: string; // JSON string for additional custom preferences
  LastModified?: Date;
  Version?: number;
}

/**
 * Dashboard Layout Configuration
 */
export interface IDashboardLayout {
  widgets: IDashboardWidget[];
  columns: number; // 1-4 columns
  rowHeight: number; // Height in pixels
  gap: number; // Gap between widgets in pixels
}

/**
 * Dashboard Widget
 */
export interface IDashboardWidget {
  id: string;
  type: WidgetType;
  title: string;
  position: {
    x: number; // Column position (0-based)
    y: number; // Row position (0-based)
    w: number; // Width in columns
    h: number; // Height in rows
  };
  settings?: any; // Widget-specific settings
  isVisible?: boolean;
  isCollapsible?: boolean;
  isCollapsed?: boolean;
  refreshInterval?: number; // Auto-refresh in seconds (0 = no refresh)
}

/**
 * Widget Types
 */
export enum WidgetType {
  MyTasks = 'my-tasks',
  MyApprovals = 'my-approvals',
  RecentProcesses = 'recent-processes',
  ProcessStats = 'process-stats',
  UpcomingDeadlines = 'upcoming-deadlines',
  TeamWorkload = 'team-workload',
  Notifications = 'notifications',
  QuickActions = 'quick-actions',
  Analytics = 'analytics',
  Calendar = 'calendar',
  AIInsights = 'ai-insights',
  CustomChart = 'custom-chart'
}

/**
 * Theme Preference
 */
export interface IThemePreference {
  mode: 'light' | 'dark' | 'auto'; // Auto follows system preference
  primaryColor?: string;
  accentColor?: string;
  useCompactMode?: boolean;
  useDepartmentBranding?: boolean;
}

/**
 * Custom Theme
 */
export interface ICustomTheme {
  id: string; // Unique theme ID
  name: string;
  description?: string;
  version?: number; // Theme version
  isActive?: boolean; // Currently active theme
  isDefault?: boolean; // System default theme
  tags?: string[]; // For categorization
  thumbnail?: string; // Base64 preview image or data URI
  createdBy?: string;
  createdDate?: Date;
  modifiedDate?: Date;

  colors: {
    // Primary brand colors
    primary?: string;
    primaryHover?: string;
    primaryPressed?: string;
    secondary?: string;
    accent?: string;

    // Semantic colors
    success?: string;
    warning?: string;
    error?: string;
    info?: string;

    // Background & surface colors
    background?: string;
    backgroundHover?: string;
    surface?: string;
    surfaceHover?: string;

    // Text colors
    text?: string;
    textSecondary?: string;
    textDisabled?: string;
    textInverted?: string;

    // Border & divider colors
    border?: string;
    borderHover?: string;
    divider?: string;

    // Link colors
    link?: string;
    linkHover?: string;
    linkPressed?: string;

    // Neutral shades (for advanced theming)
    neutral10?: string; // Lightest
    neutral20?: string;
    neutral30?: string;
    neutral40?: string;
    neutral50?: string;
    neutral60?: string;
    neutral70?: string;
    neutral80?: string;
    neutral90?: string; // Darkest
  };

  typography?: {
    fontFamily?: string;
    headingFont?: string; // Optional separate heading font
    fontSize?: number; // Base font size in pixels
    fontWeight?: number; // Base font weight
    lineHeight?: number; // Base line height
    letterSpacing?: number; // Letter spacing in pixels
  };

  spacing?: {
    baseUnit?: number; // Base spacing unit in pixels (e.g., 8)
    scale?: number[]; // Spacing scale multipliers [0.5, 1, 1.5, 2, 3, 4]
  };

  borderRadius?: {
    small?: number; // Small radius (e.g., 2px)
    medium?: number; // Medium radius (e.g., 4px)
    large?: number; // Large radius (e.g., 8px)
    round?: number; // Fully rounded (e.g., 999px)
  };

  shadows?: {
    enabled?: boolean;
    elevation1?: string; // Subtle shadow
    elevation2?: string; // Medium shadow
    elevation3?: string; // Strong shadow
    elevation4?: string; // Very strong shadow
  };

  accessibility?: {
    highContrast?: boolean;
    reducedMotion?: boolean;
    minimumTextSize?: number;
  };

  // WCAG contrast validation results
  contrastValidation?: {
    primaryOnBackground?: 'AAA' | 'AA' | 'fail';
    textOnBackground?: 'AAA' | 'AA' | 'fail';
    textOnPrimary?: 'AAA' | 'AA' | 'fail';
    validatedAt?: Date;
  };

  // Component Styles (Premium Theme Builder features)
  componentStyles?: IComponentStyles;
}

/**
 * Component Styles - Advanced styling for UI components
 */
export interface IComponentStyles {
  buttons?: IButtonStyles;
  cards?: ICardStyles;
  badges?: IBadgeStyles;
  inputs?: IInputStyles;
}

/**
 * Button Styles
 */
export interface IButtonStyles {
  // Border radius options
  borderRadius?: 'sharp' | 'rounded' | 'pill'; // sharp=0, rounded=6px, pill=999px

  // Size variants
  size?: 'compact' | 'normal' | 'large';

  // Shadow on buttons
  shadow?: 'none' | 'subtle' | 'medium' | 'strong';

  // Border width
  borderWidth?: 'none' | 'thin' | 'medium' | 'thick'; // 0, 1px, 2px, 3px

  // Animation/transition
  hoverEffect?: 'none' | 'lift' | 'glow' | 'darken';

  // Text transform
  textTransform?: 'none' | 'uppercase' | 'capitalize';

  // Font weight
  fontWeight?: 'normal' | 'medium' | 'semibold' | 'bold';
}

/**
 * Card Styles
 */
export interface ICardStyles {
  // Border radius
  borderRadius?: 'sharp' | 'small' | 'medium' | 'large'; // 0, 4px, 8px, 16px

  // Shadow intensity
  shadow?: 'none' | 'subtle' | 'medium' | 'prominent' | 'floating';

  // Border style
  borderStyle?: 'none' | 'solid' | 'dashed';
  borderWidth?: 'thin' | 'medium' | 'thick'; // 1px, 2px, 3px

  // Hover effect
  hoverEffect?: 'none' | 'lift' | 'glow' | 'border-highlight';

  // Padding
  padding?: 'compact' | 'normal' | 'spacious'; // 12px, 20px, 28px

  // Background style
  backgroundStyle?: 'solid' | 'gradient' | 'subtle-gradient';

  // Header style (for cards with headers)
  headerStyle?: 'default' | 'accent-bar' | 'full-color' | 'gradient';
}

/**
 * Badge Styles
 */
export interface IBadgeStyles {
  // Shape
  shape?: 'rounded' | 'pill' | 'square'; // 4px, 999px, 0

  // Size
  size?: 'small' | 'medium' | 'large';

  // Text style
  textTransform?: 'none' | 'uppercase' | 'capitalize';
  fontWeight?: 'normal' | 'medium' | 'semibold' | 'bold';

  // Border
  borderStyle?: 'none' | 'solid';

  // Shadow
  shadow?: 'none' | 'subtle';

  // Icon position (if badge has icon)
  iconPosition?: 'left' | 'right';
}

/**
 * Input Styles
 */
export interface IInputStyles {
  // Border radius
  borderRadius?: 'sharp' | 'small' | 'medium' | 'rounded'; // 0, 2px, 4px, 8px

  // Border style
  borderStyle?: 'full' | 'underline' | 'filled';

  // Size
  size?: 'compact' | 'normal' | 'large';

  // Focus effect
  focusEffect?: 'border' | 'glow' | 'both';

  // Label style
  labelStyle?: 'above' | 'floating' | 'inline';
}

/**
 * Saved Filter
 */
export interface ISavedFilter {
  id: string;
  name: string;
  description?: string;
  listType: 'processes' | 'tasks' | 'approvals' | 'templates';
  filters: {
    field: string;
    operator: 'eq' | 'ne' | 'gt' | 'lt' | 'ge' | 'le' | 'contains' | 'startswith' | 'endswith';
    value: any;
  }[];
  sortBy?: string;
  sortDirection?: 'asc' | 'desc';
  isDefault?: boolean;
  createdDate: Date;
  lastUsed?: Date;
}

/**
 * Saved View
 */
export interface ISavedView {
  id: string;
  name: string;
  description?: string;
  listType: 'processes' | 'tasks' | 'approvals' | 'templates';
  columns: IViewColumn[];
  filter?: ISavedFilter;
  groupBy?: string[];
  pageSize?: number;
  isDefault?: boolean;
  isShared?: boolean;
  sharedWith?: string[]; // User IDs or group names
  createdDate: Date;
  lastUsed?: Date;
}

/**
 * View Column Configuration
 */
export interface IViewColumn {
  fieldName: string;
  displayName: string;
  width?: number;
  isVisible: boolean;
  isSortable?: boolean;
  isFilterable?: boolean;
  order: number;
  format?: 'text' | 'number' | 'date' | 'currency' | 'percentage' | 'badge' | 'link';
}

/**
 * Notification Settings
 */
export interface INotificationSettings {
  email: IEmailNotificationSettings;
  inApp: IInAppNotificationSettings;
  browser?: IBrowserNotificationSettings;
  digest?: IDigestSettings;
  doNotDisturb?: IDoNotDisturbSettings;
}

/**
 * Email Notification Settings
 */
export interface IEmailNotificationSettings {
  enabled: boolean;
  events: {
    processCreated?: boolean;
    processCompleted?: boolean;
    taskAssigned?: boolean;
    taskDue?: boolean;
    taskOverdue?: boolean;
    approvalRequired?: boolean;
    approvalApproved?: boolean;
    approvalRejected?: boolean;
    processDelayed?: boolean;
    mentionedInComment?: boolean;
    delegationReceived?: boolean;
  };
  frequency?: 'immediate' | 'hourly' | 'daily' | 'weekly';
  digestTime?: string; // Time for daily/weekly digest (e.g., '09:00')
}

/**
 * In-App Notification Settings
 */
export interface IInAppNotificationSettings {
  enabled: boolean;
  events: {
    processCreated?: boolean;
    processCompleted?: boolean;
    taskAssigned?: boolean;
    taskDue?: boolean;
    taskOverdue?: boolean;
    approvalRequired?: boolean;
    approvalApproved?: boolean;
    approvalRejected?: boolean;
    processDelayed?: boolean;
    mentionedInComment?: boolean;
    delegationReceived?: boolean;
  };
  sound?: boolean;
  badge?: boolean;
  position?: 'top-right' | 'top-left' | 'bottom-right' | 'bottom-left';
  duration?: number; // Auto-dismiss after N seconds (0 = no auto-dismiss)
}

/**
 * Browser Notification Settings
 */
export interface IBrowserNotificationSettings {
  enabled: boolean;
  events: {
    approvalRequired?: boolean;
    taskDue?: boolean;
    mentionedInComment?: boolean;
  };
}

/**
 * Digest Settings
 */
export interface IDigestSettings {
  enabled: boolean;
  frequency: 'daily' | 'weekly' | 'monthly';
  dayOfWeek?: number; // 0-6 for weekly digest
  dayOfMonth?: number; // 1-31 for monthly digest
  time: string; // Time in user's timezone (e.g., '09:00')
  includeStats?: boolean;
  includePendingTasks?: boolean;
  includeUpcomingDeadlines?: boolean;
}

/**
 * Do Not Disturb Settings
 */
export interface IDoNotDisturbSettings {
  enabled: boolean;
  schedule?: {
    start: string; // Time (e.g., '22:00')
    end: string; // Time (e.g., '08:00')
    days?: number[]; // 0-6, days of week
  };
  allowUrgent?: boolean; // Allow urgent/critical notifications
}

/**
 * Department Branding
 */
export interface IDepartmentBranding extends IBaseListItem {
  Department: string;
  Logo?: string; // URL to logo image
  PrimaryColor: string;
  SecondaryColor: string;
  AccentColor: string;
  FontFamily?: string;
  CustomCSS?: string;
  IsEnabled: boolean;
}

/**
 * Favorite Item
 */
export interface IFavoriteItem {
  id: number;
  type: 'process' | 'template' | 'view' | 'filter';
  title: string;
  url?: string;
  icon?: string;
  addedDate: Date;
  lastAccessed?: Date;
  accessCount?: number;
}

/**
 * User Activity
 */
export interface IUserActivity extends IBaseListItem {
  UserId: string;
  UserEmail: string;
  ActivityType: ActivityType;
  EntityType?: string; // 'Process', 'Task', 'Approval', etc.
  EntityId?: number;
  EntityTitle?: string;
  Description?: string;
  Metadata?: string; // JSON
  IPAddress?: string;
  UserAgent?: string;
  SessionId?: string;
}

/**
 * Activity Types
 */
export enum ActivityType {
  Login = 'Login',
  Logout = 'Logout',
  ProcessViewed = 'Process Viewed',
  ProcessCreated = 'Process Created',
  ProcessUpdated = 'Process Updated',
  TaskCompleted = 'Task Completed',
  ApprovalSubmitted = 'Approval Submitted',
  PreferencesUpdated = 'Preferences Updated',
  FilterSaved = 'Filter Saved',
  ViewSaved = 'View Saved',
  TemplateUsed = 'Template Used',
  DashboardCustomized = 'Dashboard Customized',
  Search = 'Search',
  Export = 'Export',
  Import = 'Import'
}

/**
 * Keyboard Shortcut
 */
export interface IKeyboardShortcut {
  key: string; // e.g., 'Ctrl+K', 'Cmd+N'
  action: string;
  description: string;
  category: 'navigation' | 'actions' | 'editing' | 'search';
  isCustom?: boolean;
}

/**
 * Default Preferences
 */
export const DEFAULT_PREFERENCES: Partial<IUserPreferences> = {
  DashboardLayout: {
    widgets: [
      {
        id: 'my-tasks',
        type: WidgetType.MyTasks,
        title: 'My Tasks',
        position: { x: 0, y: 0, w: 2, h: 2 },
        isVisible: true
      },
      {
        id: 'my-approvals',
        type: WidgetType.MyApprovals,
        title: 'My Approvals',
        position: { x: 2, y: 0, w: 2, h: 2 },
        isVisible: true
      },
      {
        id: 'recent-processes',
        type: WidgetType.RecentProcesses,
        title: 'Recent Processes',
        position: { x: 0, y: 2, w: 2, h: 2 },
        isVisible: true
      },
      {
        id: 'quick-actions',
        type: WidgetType.QuickActions,
        title: 'Quick Actions',
        position: { x: 2, y: 2, w: 2, h: 1 },
        isVisible: true
      }
    ],
    columns: 4,
    rowHeight: 120,
    gap: 16
  },
  ThemePreference: {
    mode: 'auto',
    useCompactMode: false,
    useDepartmentBranding: true
  },
  NotificationSettings: {
    email: {
      enabled: true,
      events: {
        taskAssigned: true,
        approvalRequired: true,
        taskOverdue: true
      },
      frequency: 'immediate'
    },
    inApp: {
      enabled: true,
      events: {
        taskAssigned: true,
        approvalRequired: true,
        mentionedInComment: true
      },
      sound: true,
      badge: true,
      position: 'top-right',
      duration: 5
    },
    browser: {
      enabled: false,
      events: {
        approvalRequired: true
      }
    }
  },
  Language: 'en-US',
  TimeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  DateFormat: 'MM/DD/YYYY',
  TimeFormat: '12h',
  DensityMode: 'normal',
  DefaultItemsPerPage: 25,
  ShowWelcomeMessage: true,
  ShowTipsAndTricks: true,
  KeyboardShortcutsEnabled: true,
  AnimationsEnabled: true,
  AccessibilityMode: false,
  HighContrastMode: false
};

/**
 * Localization Support
 */
export interface ILocale {
  code: string; // 'en-US', 'es-ES', etc.
  name: string; // 'English (US)', 'Español'
  nativeName: string; // 'English (US)', 'Español'
  direction: 'ltr' | 'rtl';
  dateFormat: string;
  timeFormat: '12h' | '24h';
  numberFormat: string;
  currency: string;
}

/**
 * Default Theme Templates
 */
export const DEFAULT_THEME_TEMPLATES: ICustomTheme[] = [
  {
    id: 'microsoft-default',
    name: 'Microsoft Default',
    description: 'Default Microsoft Fluent UI theme',
    isDefault: true,
    version: 1,
    colors: {
      primary: '#0078d4',
      primaryHover: '#106ebe',
      primaryPressed: '#005a9e',
      secondary: '#106ebe',
      accent: '#8764b8',
      success: '#107c10',
      warning: '#ffc83d',
      error: '#d13438',
      info: '#0078d4',
      background: '#ffffff',
      backgroundHover: '#f3f2f1',
      surface: '#faf9f8',
      surfaceHover: '#f3f2f1',
      text: '#323130',
      textSecondary: '#605e5c',
      textDisabled: '#a19f9d',
      textInverted: '#ffffff',
      border: '#edebe9',
      borderHover: '#c8c6c4',
      divider: '#edebe9',
      link: '#0078d4',
      linkHover: '#106ebe',
      linkPressed: '#005a9e'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 4,
      large: 8,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 1.6px 3.6px rgba(0, 0, 0, 0.04)',
      elevation2: '0 3.2px 7.2px rgba(0, 0, 0, 0.08)',
      elevation3: '0 6.4px 14.4px rgba(0, 0, 0, 0.12)',
      elevation4: '0 12.8px 28.8px rgba(0, 0, 0, 0.16)'
    }
  },
  {
    id: 'dark-theme',
    name: 'Dark Mode',
    description: 'Modern dark theme optimized for low-light environments',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#3aa0f3',
      primaryHover: '#6cb8f6',
      primaryPressed: '#2886c8',
      secondary: '#2886c8',
      accent: '#a77af4',
      success: '#54b054',
      warning: '#f7b955',
      error: '#f85149',
      info: '#3aa0f3',
      background: '#1e1e1e',
      backgroundHover: '#2d2d2d',
      surface: '#252525',
      surfaceHover: '#2d2d2d',
      text: '#e1e1e1',
      textSecondary: '#b3b3b3',
      textDisabled: '#6e6e6e',
      textInverted: '#1e1e1e',
      border: '#3d3d3d',
      borderHover: '#4d4d4d',
      divider: '#3d3d3d',
      link: '#3aa0f3',
      linkHover: '#6cb8f6',
      linkPressed: '#2886c8'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 4,
      large: 8,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 1.6px 3.6px rgba(0, 0, 0, 0.32)',
      elevation2: '0 3.2px 7.2px rgba(0, 0, 0, 0.40)',
      elevation3: '0 6.4px 14.4px rgba(0, 0, 0, 0.48)',
      elevation4: '0 12.8px 28.8px rgba(0, 0, 0, 0.56)'
    }
  },
  {
    id: 'high-contrast',
    name: 'High Contrast',
    description: 'WCAG AAA compliant high contrast theme for accessibility',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#1aebff',
      primaryHover: '#6ff0ff',
      primaryPressed: '#00d4e8',
      secondary: '#00d4e8',
      accent: '#ffb900',
      success: '#00cc6a',
      warning: '#ffb900',
      error: '#ff4343',
      info: '#1aebff',
      background: '#000000',
      backgroundHover: '#1a1a1a',
      surface: '#0a0a0a',
      surfaceHover: '#1a1a1a',
      text: '#ffffff',
      textSecondary: '#e1e1e1',
      textDisabled: '#8a8a8a',
      textInverted: '#000000',
      border: '#ffffff',
      borderHover: '#e1e1e1',
      divider: '#8a8a8a',
      link: '#1aebff',
      linkHover: '#6ff0ff',
      linkPressed: '#00d4e8'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 600,
      lineHeight: 1.6,
      letterSpacing: 0.5
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 0,
      medium: 0,
      large: 0,
      round: 0
    },
    shadows: {
      enabled: false
    },
    accessibility: {
      highContrast: true,
      reducedMotion: true,
      minimumTextSize: 14
    }
  },
  {
    id: 'corporate-blue',
    name: 'Corporate Blue',
    description: 'Professional corporate theme with blue accents',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#004578',
      primaryHover: '#005a9e',
      primaryPressed: '#003150',
      secondary: '#005a9e',
      accent: '#0078d4',
      success: '#0b6a0b',
      warning: '#ca5010',
      error: '#a4262c',
      info: '#0078d4',
      background: '#f8f9fa',
      backgroundHover: '#f0f2f4',
      surface: '#ffffff',
      surfaceHover: '#f8f9fa',
      text: '#1f1f1f',
      textSecondary: '#424242',
      textDisabled: '#a6a6a6',
      textInverted: '#ffffff',
      border: '#d1d5db',
      borderHover: '#b0b7c3',
      divider: '#e5e7eb',
      link: '#004578',
      linkHover: '#005a9e',
      linkPressed: '#003150'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 3,
      medium: 6,
      large: 10,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(0, 0, 0, 0.06)',
      elevation2: '0 4px 8px rgba(0, 0, 0, 0.10)',
      elevation3: '0 8px 16px rgba(0, 0, 0, 0.14)',
      elevation4: '0 16px 32px rgba(0, 0, 0, 0.18)'
    }
  },
  {
    id: 'sunset-orange',
    name: 'Sunset Orange',
    description: 'Vibrant and energetic theme with warm orange tones',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#ff6b35',
      primaryHover: '#ff8c5a',
      primaryPressed: '#e85a25',
      secondary: '#e85a25',
      accent: '#ffa500',
      success: '#52b788',
      warning: '#ffb703',
      error: '#d62828',
      info: '#4361ee',
      background: '#fffbf7',
      backgroundHover: '#fff5ed',
      surface: '#ffffff',
      surfaceHover: '#fff5ed',
      text: '#2b2d42',
      textSecondary: '#5a5c6e',
      textDisabled: '#b5b7c5',
      textInverted: '#ffffff',
      border: '#ffd7ba',
      borderHover: '#ffb88c',
      divider: '#ffe8d6',
      link: '#ff6b35',
      linkHover: '#ff8c5a',
      linkPressed: '#e85a25'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 4,
      medium: 8,
      large: 12,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(255, 107, 53, 0.08)',
      elevation2: '0 4px 8px rgba(255, 107, 53, 0.12)',
      elevation3: '0 8px 16px rgba(255, 107, 53, 0.16)',
      elevation4: '0 16px 32px rgba(255, 107, 53, 0.20)'
    }
  },
  {
    id: 'nature-green',
    name: 'Nature Green',
    description: 'Fresh and calming theme inspired by nature',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#2d6a4f',
      primaryHover: '#40916c',
      primaryPressed: '#1b4332',
      secondary: '#40916c',
      accent: '#52b788',
      success: '#95d5b2',
      warning: '#fb8500',
      error: '#d62828',
      info: '#0077b6',
      background: '#f8fdf8',
      backgroundHover: '#f1faf2',
      surface: '#ffffff',
      surfaceHover: '#f1faf2',
      text: '#1b4332',
      textSecondary: '#2d6a4f',
      textDisabled: '#95d5b2',
      textInverted: '#ffffff',
      border: '#b7e4c7',
      borderHover: '#95d5b2',
      divider: '#d8f3dc',
      link: '#2d6a4f',
      linkHover: '#40916c',
      linkPressed: '#1b4332'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.6,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 3,
      medium: 6,
      large: 12,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(45, 106, 79, 0.06)',
      elevation2: '0 4px 8px rgba(45, 106, 79, 0.10)',
      elevation3: '0 8px 16px rgba(45, 106, 79, 0.14)',
      elevation4: '0 16px 32px rgba(45, 106, 79, 0.18)'
    }
  },
  {
    id: 'royal-purple',
    name: 'Royal Purple',
    description: 'Sophisticated and elegant theme with purple accents',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#6a4c93',
      primaryHover: '#8b6db8',
      primaryPressed: '#533a72',
      secondary: '#533a72',
      accent: '#b392c0',
      success: '#38b000',
      warning: '#ff9500',
      error: '#d74e26',
      info: '#4a90e2',
      background: '#faf8fc',
      backgroundHover: '#f5f0fa',
      surface: '#ffffff',
      surfaceHover: '#f5f0fa',
      text: '#2e1f3e',
      textSecondary: '#5a4568',
      textDisabled: '#b9a9c5',
      textInverted: '#ffffff',
      border: '#ddd0e8',
      borderHover: '#c8b4db',
      divider: '#ebe3f2',
      link: '#6a4c93',
      linkHover: '#8b6db8',
      linkPressed: '#533a72'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0.2
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 6,
      large: 10,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 6px rgba(106, 76, 147, 0.08)',
      elevation2: '0 4px 12px rgba(106, 76, 147, 0.12)',
      elevation3: '0 8px 20px rgba(106, 76, 147, 0.16)',
      elevation4: '0 16px 36px rgba(106, 76, 147, 0.20)'
    }
  },
  {
    id: 'minimalist-gray',
    name: 'Minimalist Gray',
    description: 'Clean and modern monochromatic theme',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#495057',
      primaryHover: '#6c757d',
      primaryPressed: '#343a40',
      secondary: '#6c757d',
      accent: '#868e96',
      success: '#51cf66',
      warning: '#fcc419',
      error: '#ff6b6b',
      info: '#339af0',
      background: '#ffffff',
      backgroundHover: '#f8f9fa',
      surface: '#f8f9fa',
      surfaceHover: '#e9ecef',
      text: '#212529',
      textSecondary: '#495057',
      textDisabled: '#adb5bd',
      textInverted: '#ffffff',
      border: '#dee2e6',
      borderHover: '#ced4da',
      divider: '#e9ecef',
      link: '#495057',
      linkHover: '#6c757d',
      linkPressed: '#343a40'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 4,
      large: 6,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 1px 3px rgba(0, 0, 0, 0.05)',
      elevation2: '0 2px 6px rgba(0, 0, 0, 0.08)',
      elevation3: '0 4px 12px rgba(0, 0, 0, 0.10)',
      elevation4: '0 8px 24px rgba(0, 0, 0, 0.12)'
    }
  },
  {
    id: 'warm-autumn',
    name: 'Warm Autumn',
    description: 'Cozy and inviting theme with earthy autumn colors',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#c1666b',
      primaryHover: '#d48b8f',
      primaryPressed: '#a54b50',
      secondary: '#d4a373',
      accent: '#e8b859',
      success: '#88ab75',
      warning: '#f4a261',
      error: '#d62839',
      info: '#6a8caf',
      background: '#fdf8f5',
      backgroundHover: '#f9f0e8',
      surface: '#ffffff',
      surfaceHover: '#f9f0e8',
      text: '#4a3933',
      textSecondary: '#6d5d52',
      textDisabled: '#b8a99a',
      textInverted: '#ffffff',
      border: '#e8d5c4',
      borderHover: '#d9bfa5',
      divider: '#f0e5d8',
      link: '#c1666b',
      linkHover: '#d48b8f',
      linkPressed: '#a54b50'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.6,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 4,
      medium: 8,
      large: 16,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(193, 102, 107, 0.08)',
      elevation2: '0 4px 8px rgba(193, 102, 107, 0.12)',
      elevation3: '0 8px 16px rgba(193, 102, 107, 0.16)',
      elevation4: '0 16px 32px rgba(193, 102, 107, 0.20)'
    }
  },
  {
    id: 'sky-blue',
    name: 'Sky Blue',
    description: 'Light and airy theme with soft blue tones',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#5dade2',
      primaryHover: '#85c1e9',
      primaryPressed: '#3498db',
      secondary: '#3498db',
      accent: '#a9cce3',
      success: '#58d68d',
      warning: '#f8c471',
      error: '#ec7063',
      info: '#5dade2',
      background: '#f8fbfe',
      backgroundHover: '#eef6fc',
      surface: '#ffffff',
      surfaceHover: '#eef6fc',
      text: '#1a3a52',
      textSecondary: '#4a5f7a',
      textDisabled: '#b0c4de',
      textInverted: '#ffffff',
      border: '#d0e8f5',
      borderHover: '#a9cce3',
      divider: '#e3f2fd',
      link: '#5dade2',
      linkHover: '#85c1e9',
      linkPressed: '#3498db'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 3,
      medium: 6,
      large: 12,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(93, 173, 226, 0.08)',
      elevation2: '0 4px 8px rgba(93, 173, 226, 0.12)',
      elevation3: '0 8px 16px rgba(93, 173, 226, 0.16)',
      elevation4: '0 16px 32px rgba(93, 173, 226, 0.20)'
    }
  },
  {
    id: 'olive-garden',
    name: 'Olive Garden',
    description: 'Earthy and organic theme with olive green tones',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#6b8e23',
      primaryHover: '#8fae3d',
      primaryPressed: '#556b1e',
      secondary: '#556b1e',
      accent: '#9acd32',
      success: '#90ee90',
      warning: '#daa520',
      error: '#cd5c5c',
      info: '#4682b4',
      background: '#fafaf5',
      backgroundHover: '#f4f4e8',
      surface: '#ffffff',
      surfaceHover: '#f4f4e8',
      text: '#3a3a2a',
      textSecondary: '#5a5a45',
      textDisabled: '#b8b8a0',
      textInverted: '#ffffff',
      border: '#d4d4b8',
      borderHover: '#bfbf9d',
      divider: '#e8e8d8',
      link: '#6b8e23',
      linkHover: '#8fae3d',
      linkPressed: '#556b1e'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.6,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 4,
      large: 8,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(107, 142, 35, 0.08)',
      elevation2: '0 4px 8px rgba(107, 142, 35, 0.12)',
      elevation3: '0 8px 16px rgba(107, 142, 35, 0.16)',
      elevation4: '0 16px 32px rgba(107, 142, 35, 0.20)'
    }
  },
  {
    id: 'ocean-teal',
    name: 'Ocean Teal',
    description: 'Refreshing and modern theme with teal accents',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#008080',
      primaryHover: '#20a0a0',
      primaryPressed: '#006666',
      secondary: '#006666',
      accent: '#40e0d0',
      success: '#2ecc71',
      warning: '#f39c12',
      error: '#e74c3c',
      info: '#17a2b8',
      background: '#f0fffe',
      backgroundHover: '#e0fffe',
      surface: '#ffffff',
      surfaceHover: '#e0fffe',
      text: '#1a4d4d',
      textSecondary: '#2d6a6a',
      textDisabled: '#a0c8c8',
      textInverted: '#ffffff',
      border: '#b3e5e5',
      borderHover: '#80d4d4',
      divider: '#d9f3f3',
      link: '#008080',
      linkHover: '#20a0a0',
      linkPressed: '#006666'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 4,
      medium: 8,
      large: 12,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(0, 128, 128, 0.08)',
      elevation2: '0 4px 8px rgba(0, 128, 128, 0.12)',
      elevation3: '0 8px 16px rgba(0, 128, 128, 0.16)',
      elevation4: '0 16px 32px rgba(0, 128, 128, 0.20)'
    }
  },
  {
    id: 'charcoal-steel',
    name: 'Charcoal Steel',
    description: 'Bold and industrial theme with dark charcoal tones',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#36454f',
      primaryHover: '#4f5d66',
      primaryPressed: '#262f35',
      secondary: '#2c3539',
      accent: '#5f6d78',
      success: '#4caf50',
      warning: '#ff9800',
      error: '#f44336',
      info: '#2196f3',
      background: '#fafafa',
      backgroundHover: '#f0f0f0',
      surface: '#ffffff',
      surfaceHover: '#f5f5f5',
      text: '#1a1a1a',
      textSecondary: '#424242',
      textDisabled: '#9e9e9e',
      textInverted: '#ffffff',
      border: '#d0d5d9',
      borderHover: '#b0b8bf',
      divider: '#e0e0e0',
      link: '#36454f',
      linkHover: '#4f5d66',
      linkPressed: '#262f35'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0.1
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 4,
      large: 6,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 4px rgba(54, 69, 79, 0.10)',
      elevation2: '0 4px 8px rgba(54, 69, 79, 0.14)',
      elevation3: '0 8px 16px rgba(54, 69, 79, 0.18)',
      elevation4: '0 16px 32px rgba(54, 69, 79, 0.22)'
    }
  },
  {
    id: 'burgundy-wine',
    name: 'Burgundy Wine',
    description: 'Rich and luxurious theme with deep wine red tones',
    isDefault: false,
    version: 1,
    colors: {
      primary: '#800020',
      primaryHover: '#a0193d',
      primaryPressed: '#5c0017',
      secondary: '#5c0017',
      accent: '#c9184a',
      success: '#52b788',
      warning: '#ff9f1c',
      error: '#e63946',
      info: '#3a86ff',
      background: '#fef8f9',
      backgroundHover: '#fdeef1',
      surface: '#ffffff',
      surfaceHover: '#fdeef1',
      text: '#2d1115',
      textSecondary: '#5a2a30',
      textDisabled: '#c4a5ab',
      textInverted: '#ffffff',
      border: '#e8c5cc',
      borderHover: '#d9a5b0',
      divider: '#f4dde1',
      link: '#800020',
      linkHover: '#a0193d',
      linkPressed: '#5c0017'
    },
    typography: {
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      fontSize: 14,
      fontWeight: 400,
      lineHeight: 1.5,
      letterSpacing: 0.3
    },
    spacing: {
      baseUnit: 8,
      scale: [0.5, 1, 1.5, 2, 3, 4, 6, 8]
    },
    borderRadius: {
      small: 2,
      medium: 6,
      large: 10,
      round: 999
    },
    shadows: {
      enabled: true,
      elevation1: '0 2px 6px rgba(128, 0, 32, 0.08)',
      elevation2: '0 4px 12px rgba(128, 0, 32, 0.12)',
      elevation3: '0 8px 20px rgba(128, 0, 32, 0.16)',
      elevation4: '0 16px 36px rgba(128, 0, 32, 0.20)'
    }
  }
];

/**
 * Supported Locales
 */
export const SUPPORTED_LOCALES: ILocale[] = [
  {
    code: 'en-US',
    name: 'English (US)',
    nativeName: 'English (US)',
    direction: 'ltr',
    dateFormat: 'MM/DD/YYYY',
    timeFormat: '12h',
    numberFormat: 'en-US',
    currency: 'USD'
  },
  {
    code: 'en-GB',
    name: 'English (UK)',
    nativeName: 'English (UK)',
    direction: 'ltr',
    dateFormat: 'DD/MM/YYYY',
    timeFormat: '24h',
    numberFormat: 'en-GB',
    currency: 'GBP'
  },
  {
    code: 'es-ES',
    name: 'Spanish',
    nativeName: 'Español',
    direction: 'ltr',
    dateFormat: 'DD/MM/YYYY',
    timeFormat: '24h',
    numberFormat: 'es-ES',
    currency: 'EUR'
  },
  {
    code: 'fr-FR',
    name: 'French',
    nativeName: 'Français',
    direction: 'ltr',
    dateFormat: 'DD/MM/YYYY',
    timeFormat: '24h',
    numberFormat: 'fr-FR',
    currency: 'EUR'
  },
  {
    code: 'de-DE',
    name: 'German',
    nativeName: 'Deutsch',
    direction: 'ltr',
    dateFormat: 'DD.MM.YYYY',
    timeFormat: '24h',
    numberFormat: 'de-DE',
    currency: 'EUR'
  },
  {
    code: 'pt-BR',
    name: 'Portuguese (Brazil)',
    nativeName: 'Português (Brasil)',
    direction: 'ltr',
    dateFormat: 'DD/MM/YYYY',
    timeFormat: '24h',
    numberFormat: 'pt-BR',
    currency: 'BRL'
  },
  {
    code: 'ja-JP',
    name: 'Japanese',
    nativeName: '日本語',
    direction: 'ltr',
    dateFormat: 'YYYY/MM/DD',
    timeFormat: '24h',
    numberFormat: 'ja-JP',
    currency: 'JPY'
  },
  {
    code: 'zh-CN',
    name: 'Chinese (Simplified)',
    nativeName: '简体中文',
    direction: 'ltr',
    dateFormat: 'YYYY-MM-DD',
    timeFormat: '24h',
    numberFormat: 'zh-CN',
    currency: 'CNY'
  }
];
