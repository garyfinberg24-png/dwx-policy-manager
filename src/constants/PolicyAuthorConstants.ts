/**
 * Constants used by the PolicyAuthorEnhanced component.
 * Centralizes magic numbers and repeated string values.
 */

/** Auto-save interval in milliseconds (1 minute) */
export const AUTO_SAVE_INTERVAL_MS = 60000;

/** Default page sizes for SharePoint list queries */
export const QUERY_PAGE_SIZE = {
  STANDARD: 100,
  SMALL: 50,
  LARGE: 200,
  MAX: 500,
} as const;

/** URL parameter names used by the Policy Builder */
export const URL_PARAMS = {
  EDIT_POLICY_ID: 'editPolicyId',
  TAB: 'tab',
  QUIZ_ID: 'quizId',
} as const;

/** Default values for new policies */
export const POLICY_DEFAULTS = {
  READ_TIMEFRAME_DAYS: 7,
  VERSION: '1.0',
  REVIEW_FREQUENCY: 'Annually',
} as const;

/** PeoplePicker configuration */
export const PEOPLE_PICKER = {
  MAX_REVIEWERS: 10,
  MAX_APPROVERS: 5,
  RESOLVE_DELAY_MS: 1000,
} as const;

/** Bulk import policy number prefix */
export const BULK_IMPORT_PREFIX = 'POL-IMP';
