/**
 * Fluent UI Shared Styles for JML SharePoint Solution
 *
 * This file contains shared style definitions following Microsoft Fluent UI Design System
 * https://developer.microsoft.com/en-us/fluentui
 */

export const FluentColors = {
  // Primary Colors
  themePrimary: '#0078d4',
  themeLighterAlt: '#eff6fc',
  themeLighter: '#deecf9',
  themeLight: '#c7e0f4',
  themeTertiary: '#71afe5',
  themeSecondary: '#2b88d8',
  themeDarkAlt: '#106ebe',
  themeDark: '#005a9e',
  themeDarker: '#004578',

  // Neutral Colors
  neutralLighterAlt: '#faf9f8',
  neutralLighter: '#f3f2f1',
  neutralLight: '#edebe9',
  neutralQuaternaryAlt: '#e1dfdd',
  neutralQuaternary: '#d2d0ce',
  neutralTertiaryAlt: '#c8c6c4',
  neutralTertiary: '#a19f9d',
  neutralSecondary: '#605e5c',
  neutralPrimaryAlt: '#3b3a39',
  neutralPrimary: '#323130',
  neutralDark: '#201f1e',
  black: '#000000',
  white: '#ffffff',

  // Semantic Colors
  success: '#107c10',
  successLight: '#dff6dd',
  warning: '#fde300',
  warningDark: '#8a6116',
  warningLight: '#fff4ce',
  error: '#a80000',
  errorLight: '#fde7e9',
  info: '#0078d4',
  infoLight: '#e1f5fe',

  // Action Colors
  actionTeal: '#03787C',
  actionTealHover: '#026569',
  actionTealPressed: '#015256',
};

export const FluentTypography = {
  fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',

  // Standardized Typography Scale (Fluent UI 2 spec)
  sizes: {
    hero: '42px',      // Rarely used - hero sections
    title: '28px',     // H2 - Section headers
    subtitle: '20px',  // H3 - Subsection headers
    large: '18px',     // Emphasized text
    medium: '16px',    // H4 - Card titles, large body
    normal: '14px',    // Default body text
    small: '12px',     // Caption, helper text
    tiny: '10px',      // Very small labels
  },
  weights: {
    light: '300',
    regular: '400',
    semibold: '500',
    bold: '600',
  },

  // Shorthand aliases for convenience (as objects to match h1-h4 structure)
  light: { fontWeight: '300' },
  regular: { fontWeight: '400' },
  semibold: { fontWeight: '500' },
  bold: { fontWeight: '600' },

  // Semantic Typography Helpers (H1-H4 hierarchy)
  h1: {
    fontSize: '32px',
    fontWeight: '600',
    lineHeight: '40px',
  },
  h2: {
    fontSize: '24px',
    fontWeight: '600',
    lineHeight: '32px',
  },
  h3: {
    fontSize: '20px',
    fontWeight: '600',
    lineHeight: '28px',
  },
  h4: {
    fontSize: '16px',
    fontWeight: '600',
    lineHeight: '22px',
  },
  bodyLarge: {
    fontSize: '16px',
    fontWeight: '400',
    lineHeight: '22px',
  },
  body: {
    fontSize: '14px',
    fontWeight: '400',
    lineHeight: '20px',
  },
  caption: {
    fontSize: '12px',
    fontWeight: '400',
    lineHeight: '16px',
  },
  badge: {
    fontSize: '12px',
    fontWeight: '600',
    lineHeight: '16px',
  },
};

export const FluentSpacing = {
  xxs: '4px',  // Alias for xs
  xs: '4px',
  s: '8px',
  m: '12px',
  l: '16px',
  xl: '20px',
  xxl: '24px',
  xxxl: '32px',
};

export const FluentBorderRadius = {
  small: '2px',
  medium: '4px',
  large: '8px',      // Standard for cards
  round: '50%',
  button: '6px',     // JML Standard for buttons
};

/**
 * Icon Size Standards
 */
export const IconSizes = {
  small: '16px',      // Inline with text, badges
  medium: '20px',     // Buttons, list items
  large: '24px',      // Page headers, primary actions
  xlarge: '32px',     // Stat cards, empty states
  hero: '48px',       // Large empty states, splash screens
};

/**
 * Icon Color Standards
 */
export const IconColors = {
  primary: '#0078d4',         // Primary actions, highlights
  success: '#107c10',         // Success states, completed
  warning: '#f7630c',         // Warning states, attention
  error: '#d13438',           // Error states, critical
  neutral: '#605e5c',         // Default, neutral actions
  neutralLight: '#8a8886',    // Disabled, subtle
  white: '#ffffff',           // On colored backgrounds
};

export const FluentShadows = {
  depth4: '0 1.6px 3.6px 0 rgba(0,0,0,.132), 0 0.3px 0.9px 0 rgba(0,0,0,.108)',
  depth8: '0 3.2px 7.2px 0 rgba(0,0,0,.132), 0 0.6px 1.8px 0 rgba(0,0,0,.108)',
  depth16: '0 6.4px 14.4px 0 rgba(0,0,0,.182), 0 1.2px 3.6px 0 rgba(0,0,0,.148)',
  depth64: '0 25.6px 57.6px 0 rgba(0,0,0,.22), 0 4.8px 14.4px 0 rgba(0,0,0,.18)',
};

export const FluentAnimations = {
  durationFast: '0.1s',
  durationNormal: '0.2s',
  durationSlow: '0.4s',
  easeIn: 'cubic-bezier(0.4, 0, 1, 1)',
  easeOut: 'cubic-bezier(0, 0, 0.2, 1)',
  easeInOut: 'cubic-bezier(0.4, 0, 0.2, 1)',
  easingDefault: 'cubic-bezier(0.4, 0, 0.2, 1)',  // Alias for easeInOut
};

/**
 * Get badge style based on type
 */
export const getBadgeStyle = (type: 'success' | 'warning' | 'error' | 'info' | 'subtle') => {
  const baseStyle = {
    display: 'inline-flex',
    alignItems: 'center',
    padding: `${FluentSpacing.xs} ${FluentSpacing.m}`,
    borderRadius: FluentBorderRadius.large,
    fontSize: FluentTypography.sizes.small,
    fontWeight: FluentTypography.weights.semibold,
  };

  switch (type) {
    case 'success':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.successLight,
        color: FluentColors.success,
        border: `1px solid ${FluentColors.success}`,
      };
    case 'warning':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.warningLight,
        color: FluentColors.warningDark,
        border: `1px solid ${FluentColors.warning}`,
      };
    case 'error':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.errorLight,
        color: FluentColors.error,
        border: `1px solid ${FluentColors.error}`,
      };
    case 'info':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.infoLight,
        color: FluentColors.info,
        border: `1px solid ${FluentColors.info}`,
      };
    case 'subtle':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.neutralLighter,
        color: FluentColors.neutralSecondary,
        border: `1px solid ${FluentColors.neutralQuaternary}`,
      };
    default:
      return baseStyle;
  }
};

/**
 * Get button style based on type (Three-Tier Button System)
 * PRIMARY = Main CTAs (Save, Submit, Create)
 * SECONDARY/DEFAULT = Cancel, Back, Alternative actions
 * TERTIARY/SUBTLE = More actions, filters, info
 */
export const getButtonStyle = (type: 'primary' | 'secondary' | 'default' | 'tertiary' | 'subtle' = 'default') => {
  const baseStyle = {
    display: 'inline-flex',
    alignItems: 'center',
    gap: FluentSpacing.s,
    padding: '10px 20px',                    // Standardized padding
    cursor: 'pointer',
    borderRadius: FluentBorderRadius.medium, // 4px for all buttons
    fontSize: FluentTypography.sizes.normal,
    fontWeight: FluentTypography.weights.semibold,
    fontFamily: FluentTypography.fontFamily,
    transition: `all ${FluentAnimations.durationNormal} ${FluentAnimations.easeInOut}`,
    border: 'none',
    ':active': {
      transform: 'scale(0.98)',              // Press effect
    },
  };

  switch (type) {
    case 'primary':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.themePrimary,  // Microsoft Blue (#0078d4)
        color: FluentColors.white,
        border: 'none',
        ':hover': {
          backgroundColor: FluentColors.themeDarkAlt,
          transform: 'translateY(-1px)',              // Subtle lift
          boxShadow: '0 4px 8px rgba(0,0,0,0.15)',
        },
        ':active': {
          transform: 'scale(0.98)',
          backgroundColor: FluentColors.themeDark,
        },
      };

    case 'secondary':
    case 'default':
      return {
        ...baseStyle,
        backgroundColor: FluentColors.white,
        color: FluentColors.neutralPrimary,
        border: `1px solid ${FluentColors.neutralTertiaryAlt}`,
        ':hover': {
          backgroundColor: FluentColors.neutralLighter,
          borderColor: FluentColors.neutralPrimary,
        },
        ':active': {
          transform: 'scale(0.98)',
          backgroundColor: FluentColors.neutralLight,
        },
      };

    case 'tertiary':
    case 'subtle':
      return {
        ...baseStyle,
        backgroundColor: 'transparent',
        color: FluentColors.neutralSecondary,
        border: '1px solid transparent',
        padding: '8px 12px',                          // Smaller padding for tertiary
        ':hover': {
          backgroundColor: FluentColors.neutralLighter,
          color: FluentColors.neutralPrimary,
        },
        ':active': {
          transform: 'scale(0.98)',
        },
      };

    default:
      return {
        ...baseStyle,
        backgroundColor: FluentColors.white,
        color: FluentColors.neutralPrimary,
        border: `1px solid ${FluentColors.neutralTertiaryAlt}`,
        ':hover': {
          backgroundColor: FluentColors.neutralLighter,
          borderColor: FluentColors.neutralPrimary,
        },
      };
  }
};

/**
 * Get standardized card style
 * All cards MUST use: 8px radius, depth4 shadow, subtle border
 */
export const getCardStyle = (elevated: boolean = true, interactive: boolean = false) => ({
  backgroundColor: FluentColors.white,
  borderRadius: FluentBorderRadius.large,        // 8px - enforced standard
  border: `1px solid ${FluentColors.neutralLight}`,
  padding: FluentSpacing.xl,                     // 20px - standard card padding
  boxShadow: elevated ? FluentShadows.depth4 : 'none',
  transition: `all ${FluentAnimations.durationNormal} ${FluentAnimations.easeInOut}`,

  // Interactive cards (clickable) have hover effects
  ...(interactive && {
    cursor: 'pointer',
    ':hover': {
      transform: 'translateY(-2px)',
      boxShadow: FluentShadows.depth16,
      borderColor: FluentColors.themePrimary,    // Blue highlight on hover
    },
  }),
});

/**
 * Get progress bar style
 */
export const getProgressBarStyle = () => ({
  bar: {
    width: '100%',
    height: '6px',
    backgroundColor: FluentColors.neutralLight,
    borderRadius: FluentBorderRadius.small,
    overflow: 'hidden',
  },
  fill: {
    height: '100%',
    backgroundColor: FluentColors.themePrimary,
    borderRadius: FluentBorderRadius.small,
    transition: `width ${FluentAnimations.durationNormal} ${FluentAnimations.easeOut}`,
  },
});

/**
 * Get standardized header styles for full-width webparts
 * Matches the design from Survey Management and CV Management
 */
export const getStandardHeaderStyles = () => ({
  container: {
    background: 'linear-gradient(135deg, #0078d4 0%, #106ebe 100%)',
    padding: FluentSpacing.xxl,
    marginBottom: FluentSpacing.xl,
    position: 'relative' as const,
    overflow: 'hidden',
    '::before': {
      content: '""',
      position: 'absolute' as const,
      top: 0,
      right: 0,
      bottom: 0,
      left: 0,
      background: 'radial-gradient(circle at top right, rgba(255,255,255,0.1) 0%, transparent 60%)',
      pointerEvents: 'none' as const,
    },
  },
  contentRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    position: 'relative' as const,
    zIndex: 1,
  },
  titleSection: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
  },
  icon: {
    fontSize: '24px',
    color: FluentColors.white,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  titleBlock: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.xs,
  },
  title: {
    fontSize: FluentTypography.sizes.subtitle,
    fontWeight: FluentTypography.weights.semibold,
    color: FluentColors.white,
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },
  description: {
    fontSize: FluentTypography.sizes.small,
    color: 'rgba(255, 255, 255, 0.9)',
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },
  actionButtons: {
    display: 'flex',
    gap: FluentSpacing.m,
    alignItems: 'center',
  },
});

/**
 * Get action button style (teal color for primary actions)
 */
export const getActionButtonStyle = () => ({
  backgroundColor: FluentColors.actionTeal,
  color: FluentColors.white,
  border: `1px solid ${FluentColors.actionTeal}`,
  borderRadius: FluentBorderRadius.small,
  padding: `${FluentSpacing.s} ${FluentSpacing.l}`,
  fontSize: FluentTypography.sizes.normal,
  fontWeight: FluentTypography.weights.semibold,
  fontFamily: FluentTypography.fontFamily,
  cursor: 'pointer',
  transition: `all ${FluentAnimations.durationNormal} ${FluentAnimations.easeInOut}`,
  display: 'inline-flex',
  alignItems: 'center',
  gap: FluentSpacing.s,
  ':hover': {
    backgroundColor: FluentColors.actionTealHover,
    borderColor: FluentColors.actionTealHover,
  },
  ':active': {
    backgroundColor: FluentColors.actionTealPressed,
    borderColor: FluentColors.actionTealPressed,
    transform: 'scale(0.98)',
  },
});

/**
 * Get filter row layout styles (horizontal inline layout)
 * Matches the CV Management screen layout
 */
export const getFilterRowStyles = () => ({
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
    marginBottom: FluentSpacing.l,
    flexWrap: 'wrap' as const,
  },
  searchField: {
    flex: '1 1 300px',
    minWidth: '200px',
  },
  dropdown: {
    flex: '0 0 auto',
    minWidth: '150px',
  },
  actionGroup: {
    marginLeft: 'auto',
    display: 'flex',
    gap: FluentSpacing.s,
    alignItems: 'center',
  },
});

/**
 * Get statistics card grid styles (horizontal layout)
 */
export const getStatsCardGridStyles = () => ({
  container: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
    gap: FluentSpacing.l,
    marginBottom: FluentSpacing.xl,
  },
  card: {
    ...getCardStyle(true),
    padding: FluentSpacing.l,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.s,
    ':hover': {
      boxShadow: FluentShadows.depth8,
      transform: 'translateY(-2px)',
    },
  },
  iconRow: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.s,
  },
  icon: {
    fontSize: '20px',
    color: FluentColors.themePrimary,
  },
  label: {
    fontSize: FluentTypography.sizes.small,
    color: FluentColors.neutralSecondary,
    fontFamily: FluentTypography.fontFamily,
  },
  value: {
    fontSize: FluentTypography.sizes.title,
    fontWeight: FluentTypography.weights.semibold,
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },
  subtext: {
    fontSize: FluentTypography.sizes.small,
    color: FluentColors.neutralTertiary,
    fontFamily: FluentTypography.fontFamily,
  },
});

/**
 * Get full-width content container styles
 */
export const getFullWidthContainerStyles = () => ({
  width: '100%',
  maxWidth: '100%',
  padding: 0,
  margin: 0,
});

/**
 * Get full-width table styles
 */
export const getFullWidthTableStyles = () => ({
  width: '100%',
  borderCollapse: 'collapse' as const,
  fontFamily: FluentTypography.fontFamily,
  fontSize: FluentTypography.sizes.normal,
});

/**
 * Get standardized table component styles
 * All data tables MUST use Fluent UI v9 Table component with these styles
 */
export const getStandardTableStyles = () => ({
  table: {
    width: '100%',
    borderCollapse: 'collapse' as const,
    fontFamily: FluentTypography.fontFamily,
    fontSize: FluentTypography.sizes.normal,
    backgroundColor: FluentColors.white,
    borderRadius: FluentBorderRadius.large,      // 8px
    overflow: 'hidden',
    boxShadow: FluentShadows.depth4,
  },

  headerRow: {
    backgroundColor: FluentColors.neutralLighter,  // #f3f2f1
    borderBottom: `2px solid ${FluentColors.neutralLight}`,
  },

  headerCell: {
    color: FluentColors.neutralPrimary,
    fontSize: FluentTypography.sizes.normal,
    fontWeight: FluentTypography.weights.semibold,
    padding: '12px 16px',                         // Standardized cell padding
    textAlign: 'left' as const,
    borderBottom: `2px solid ${FluentColors.neutralLight}`,
  },

  row: {
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
    transition: `background-color ${FluentAnimations.durationFast} ${FluentAnimations.easeOut}`,
    ':hover': {
      backgroundColor: FluentColors.neutralLighter,  // Subtle row highlight
      cursor: 'pointer',
    },
  },

  cell: {
    padding: '12px 16px',                         // Standardized cell padding
    color: FluentColors.neutralPrimary,
    fontSize: FluentTypography.sizes.normal,
    verticalAlign: 'middle' as const,
  },

  emptyRow: {
    textAlign: 'center' as const,
    padding: FluentSpacing.xxl,
    color: FluentColors.neutralTertiary,
  },
});

/**
 * Get stat card styles (for Dashboard KPIs)
 */
export const getStatCardStyle = () => ({
  ...getCardStyle(true, true),                    // Elevated + Interactive
  minHeight: '120px',
  display: 'flex',
  flexDirection: 'column' as const,
  alignItems: 'center',
  justifyContent: 'center',
  gap: FluentSpacing.s,
  cursor: 'pointer',
});

/**
 * Get standardized page container styles
 */
export const getPageContainerStyles = () => ({
  display: 'flex',
  flexDirection: 'column' as const,
  gap: FluentSpacing.xxl,                         // 24px
  padding: FluentSpacing.xxl,                     // 24px page padding
  backgroundColor: FluentColors.neutralLighterAlt, // #faf9f8
  minHeight: '100vh',
  fontFamily: FluentTypography.fontFamily,
});
