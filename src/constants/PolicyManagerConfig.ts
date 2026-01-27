// ============================================================================
// DWx Policy Manager - Application Configuration
// Centralized configuration for the Policy Manager application
// ============================================================================

/**
 * SharePoint Site Configuration
 */
export const SiteConfig = {
  /** Policy Manager SharePoint site URL */
  SITE_URL: 'https://mf7m.sharepoint.com/sites/PolicyManager',

  /** Tenant ID */
  TENANT_ID: '03bbbdee-d78b-4613-9b99-c468398246b7',

  /** Application name */
  APP_NAME: 'DWx Policy Manager',

  /** Application short name */
  APP_SHORT_NAME: 'Policy Manager',

  /** Version */
  VERSION: '1.0.0',
} as const;

/**
 * Feature Flags
 * Enable/disable features for gradual rollout
 */
export const FeatureFlags = {
  /** Enable quiz functionality */
  ENABLE_QUIZZES: true,

  /** Enable social features (ratings, comments) */
  ENABLE_SOCIAL: true,

  /** Enable policy packs */
  ENABLE_POLICY_PACKS: true,

  /** Enable gamification/badges */
  ENABLE_GAMIFICATION: true,

  /** Enable advanced analytics */
  ENABLE_ANALYTICS: true,

  /** Enable document attachments */
  ENABLE_DOCUMENTS: true,

  /** Enable email notifications */
  ENABLE_NOTIFICATIONS: true,

  /** Enable Teams integration */
  ENABLE_TEAMS_INTEGRATION: true,
} as const;

/**
 * Default Values
 * Application-wide default settings
 */
export const Defaults = {
  /** Default items per page for lists */
  ITEMS_PER_PAGE: 20,

  /** Default acknowledgement deadline in days */
  ACKNOWLEDGEMENT_DEADLINE_DAYS: 14,

  /** Default quiz passing score percentage */
  QUIZ_PASSING_SCORE: 80,

  /** Default review cycle in months */
  REVIEW_CYCLE_MONTHS: 12,

  /** Cache duration in minutes */
  CACHE_DURATION_MINUTES: 15,

  /** Maximum file upload size in MB */
  MAX_FILE_SIZE_MB: 25,
} as const;

/**
 * DWx Brand Colors
 * Based on DWx Brand Guidelines
 */
export const BrandColors = {
  /** Primary Blue */
  PRIMARY: '#1a5a8a',

  /** Primary Dark */
  PRIMARY_DARK: '#0d3a5c',

  /** Primary Light */
  PRIMARY_LIGHT: '#2d7ab8',

  /** Primary gradient */
  PRIMARY_GRADIENT: 'linear-gradient(135deg, #1a5a8a 0%, #2d7ab8 100%)',

  /** Success Green */
  SUCCESS: '#107c10',

  /** Warning Orange */
  WARNING: '#ff8c00',

  /** Error Red */
  ERROR: '#d13438',

  /** Info Blue */
  INFO: '#0078d4',
} as const;

/**
 * API Endpoints
 * External service endpoints
 */
export const ApiEndpoints = {
  /** Microsoft Graph API base URL */
  GRAPH_API: 'https://graph.microsoft.com/v1.0',

  /** SharePoint REST API path */
  SP_REST_API: '/_api',
} as const;

// Export all configuration
export const PolicyManagerConfig = {
  Site: SiteConfig,
  Features: FeatureFlags,
  Defaults,
  Colors: BrandColors,
  Api: ApiEndpoints,
} as const;

export default PolicyManagerConfig;
