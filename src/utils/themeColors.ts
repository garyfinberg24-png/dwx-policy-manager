/**
 * themeColors — Runtime theme color helper for inline React styles.
 *
 * CSS variables (var(--pm-primary)) work in SCSS but NOT in React inline styles.
 * This helper reads the computed CSS variable values and returns them as strings
 * for use in style={{ color: tc.primary }} patterns.
 *
 * Performance: getComputedStyle() is called once per variable, then cached.
 * Cache is invalidated when ThemeManager.apply() fires (via mutation observer).
 *
 * Usage:
 *   import { tc } from '../../utils/themeColors';
 *   <div style={{ color: tc.primary, background: tc.headerBg }}>
 */

const DEFAULTS: Record<string, string> = {
  '--pm-primary': '#0d9488',
  '--pm-primary-dark': '#0f766e',
  '--pm-primary-light': '#ccfbf1',
  '--pm-primary-lighter': '#f0fdfa',
  '--pm-header-bg': 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
  '--pm-accent': '#2563eb',
  '--pm-success': '#059669',
  '--pm-warning': '#d97706',
  '--pm-danger': '#dc2626',
  '--pm-card-bg': '#ffffff',
  '--pm-content-bg': '#f1f5f9',
  '--pm-sidebar-bg': '#f8fafc',
  '--pm-card-radius': '10px',
  '--pm-control-radius': '4px',
  '--pm-font-family': "'Segoe UI', -apple-system, sans-serif",
};

let cache: Record<string, string> = {};
let cacheValid = false;

function readVar(varName: string): string {
  if (typeof document === 'undefined') return DEFAULTS[varName] || '';
  const val = getComputedStyle(document.documentElement).getPropertyValue(varName).trim();
  return val || DEFAULTS[varName] || '';
}

function buildCache(): void {
  cache = {};
  for (const key of Object.keys(DEFAULTS)) {
    cache[key] = readVar(key);
  }
  cacheValid = true;
}

/** Invalidate cache — called by ThemeManager after applying a new theme */
export function invalidateThemeCache(): void {
  cacheValid = false;
}

/** Get a theme color value for use in inline styles */
function get(varName: string): string {
  if (!cacheValid) buildCache();
  return cache[varName] || DEFAULTS[varName] || '';
}

/**
 * Theme colors object — use in inline styles.
 * Access is lazy (reads on first use per render cycle).
 *
 * Example: style={{ color: tc.primary, background: tc.headerBg }}
 */
export const tc = {
  get primary() { return get('--pm-primary'); },
  get primaryDark() { return get('--pm-primary-dark'); },
  get primaryLight() { return get('--pm-primary-light'); },
  get primaryLighter() { return get('--pm-primary-lighter'); },
  get headerBg() { return get('--pm-header-bg'); },
  get accent() { return get('--pm-accent'); },
  get success() { return get('--pm-success'); },
  get warning() { return get('--pm-warning'); },
  get danger() { return get('--pm-danger'); },
  get cardBg() { return get('--pm-card-bg'); },
  get contentBg() { return get('--pm-content-bg'); },
  get sidebarBg() { return get('--pm-sidebar-bg'); },
  get cardRadius() { return get('--pm-card-radius'); },
  get controlRadius() { return get('--pm-control-radius'); },
  get fontFamily() { return get('--pm-font-family'); },
};
