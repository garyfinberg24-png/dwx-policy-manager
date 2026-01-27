/**
 * JML Solution - Centralized Design Tokens
 * Based on Fluent UI Design System and SharePoint Theme
 *
 * This file provides consistent design tokens across all JML web parts
 * following the standards defined in .claude/instructions.md
 *
 * @version 1.0.0
 * @see https://developer.microsoft.com/en-us/fluentui
 */

/**
 * Color Tokens
 * Based on SharePoint theme colors and Fluent UI semantic colors
 */
export const JMLColors = {
  // Primary Brand Colors
  themePrimary: '#0078d4',
  themeDark: '#106ebe',
  themeDarker: '#005a9e',
  themeLight: '#c7e0f4',
  themeLighter: '#deecf9',
  themeLighterAlt: '#eff6fc',

  // Semantic Colors
  success: '#4caf50',
  successLight: '#e8f5e9',
  warning: '#f57c00',
  warningLight: '#fff3e0',
  error: '#f44336',
  errorLight: '#ffebee',
  info: '#0078d4',
  infoLight: '#e3f2fd',

  // Neutral Colors
  neutralPrimary: '#323130',
  neutralSecondary: '#605e5c',
  neutralTertiary: '#a19f9d',
  neutralLight: '#edebe9',
  neutralLighter: '#f3f2f1',
  neutralLighterAlt: '#faf9f8',
  white: '#ffffff',

  // Status Colors (for badges and indicators)
  statusJoiner: '#0078d4',
  statusMover: '#9c27b0',
  statusLeaver: '#f57c00',
  statusCompleted: '#4caf50',
  statusInProgress: '#0078d4',
  statusPending: '#f57c00',
  statusOverdue: '#f44336',
  statusCancelled: '#9e9e9e',
  statusOnHold: '#fdd835',

  // Priority Colors
  priorityHigh: '#f44336',
  priorityMedium: '#f57c00',
  priorityLow: '#4caf50',
  priorityCritical: '#d32f2f',
} as const;

/**
 * Typography Tokens
 * Font sizes, weights, and line heights
 */
export const JMLTypography = {
  // Font Family
  fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",

  // Font Sizes
  fontSize: {
    /** 32px - Page hero titles */
    hero: '32px',
    /** 28px - Main page titles (STANDARD) */
    title: '28px',
    /** 24px - Large headings */
    h1: '24px',
    /** 20px - Section headings */
    h2: '20px',
    /** 18px - Subsection headings */
    h3: '18px',
    /** 16px - Card titles, emphasized text */
    h4: '16px',
    /** 14px - Body text, labels */
    base: '14px',
    /** 13px - Secondary text */
    small: '13px',
    /** 12px - Caption text, metadata */
    caption: '12px',
    /** 11px - Tiny text, legal */
    tiny: '11px',
  },

  // Font Weights
  fontWeight: {
    regular: 400,
    medium: 500,
    semibold: 600,
    bold: 700,
  },

  // Line Heights
  lineHeight: {
    tight: 1.2,
    normal: 1.5,
    relaxed: 1.75,
  },
} as const;

/**
 * Spacing Tokens
 * Consistent spacing scale based on 4px grid
 */
export const JMLSpacing = {
  /** 4px */
  xxs: '4px',
  /** 8px */
  xs: '8px',
  /** 12px */
  s: '12px',
  /** 16px */
  m: '16px',
  /** 20px */
  l: '20px',
  /** 24px */
  xl: '24px',
  /** 32px */
  xxl: '32px',
  /** 40px */
  xxxl: '40px',
  /** 48px */
  huge: '48px',
  /** 64px */
  massive: '64px',
} as const;

/**
 * Border Radius Tokens
 * Standard corner rounding values
 */
export const JMLBorderRadius = {
  /** 2px - Minimal rounding */
  small: '2px',
  /** 4px - Standard buttons, inputs */
  medium: '4px',
  /** 8px - Cards, containers (STANDARD) */
  large: '8px',
  /** 12px - Badges, pills */
  xlarge: '12px',
  /** 50% - Circular avatars */
  circle: '50%',
} as const;

/**
 * Shadow Tokens (Elevation)
 * Based on Fluent UI shadow system
 */
export const JMLShadows = {
  /** No shadow */
  none: 'none',
  /** 0 2px 4px rgba(0, 0, 0, 0.08) - Cards, panels (STANDARD) */
  depth4: '0 2px 4px rgba(0, 0, 0, 0.08)',
  /** 0 4px 8px rgba(0, 0, 0, 0.12) - Hover states, emphasis */
  depth8: '0 4px 8px rgba(0, 0, 0, 0.12)',
  /** 0 8px 16px rgba(0, 0, 0, 0.14) - Elevated cards */
  depth16: '0 8px 16px rgba(0, 0, 0, 0.14)',
  /** 0 16px 32px rgba(0, 0, 0, 0.18) - Modals, flyouts */
  depth64: '0 16px 32px rgba(0, 0, 0, 0.18)',
} as const;

/**
 * Header Tokens
 * Consistent header styling across all web parts
 */
export const JMLHeader = {
  /** Gradient background */
  background: 'linear-gradient(135deg, #0078d4 0%, #106ebe 100%)',
  /** Text color */
  color: '#ffffff',
  /** Padding (desktop) */
  padding: '24px 40px',
  /** Padding (mobile) */
  paddingMobile: '16px 20px',
  /** Shadow */
  shadow: '0 2px 8px rgba(0, 0, 0, 0.1)',
  /** Bottom margin */
  marginBottom: '32px',

  /** Icon size (STANDARD) */
  iconSize: '32px',
  /** Title font size (STANDARD) */
  titleSize: '28px',
  /** Title font weight */
  titleWeight: 600,
  /** Subtitle font size */
  subtitleSize: '14px',
  /** Subtitle opacity */
  subtitleOpacity: 0.9,
} as const;

/**
 * Card Tokens
 * Standard card component styling
 */
export const JMLCard = {
  /** Background color */
  background: '#ffffff',
  /** Border color */
  borderColor: '#edebe9',
  /** Border width */
  borderWidth: '1px',
  /** Border radius (STANDARD: 8px) */
  borderRadius: '8px',
  /** Padding */
  padding: '20px',
  /** Shadow (default) */
  shadow: '0 2px 4px rgba(0, 0, 0, 0.08)',
  /** Shadow (hover) */
  shadowHover: '0 4px 8px rgba(0, 0, 0, 0.12)',
  /** Hover border color */
  borderColorHover: '#0078d4',
  /** Hover transform */
  transformHover: 'translateY(-2px)',
  /** Transition duration */
  transition: 'all 0.2s ease',
} as const;

/**
 * Container Tokens
 * Main content container styling
 */
export const JMLContainer = {
  /** Maximum width */
  maxWidth: '1600px',
  /** Padding (desktop) */
  padding: '0 40px 40px',
  /** Padding (mobile) */
  paddingMobile: '0 20px 20px',
  /** Background color */
  background: '#faf9f8',
} as const;

/**
 * Button Tokens
 * Standard button styling
 */
export const JMLButton = {
  // Primary Button
  primary: {
    background: '#0078d4',
    backgroundHover: '#106ebe',
    color: '#ffffff',
    borderRadius: '4px',
    padding: '8px 16px',
    fontSize: '14px',
    fontWeight: 600,
  },

  // Secondary Button
  secondary: {
    background: 'transparent',
    backgroundHover: '#f3f2f1',
    color: '#0078d4',
    border: '1px solid #0078d4',
    borderRadius: '4px',
    padding: '8px 16px',
    fontSize: '14px',
    fontWeight: 600,
  },

  // Subtle Button
  subtle: {
    background: 'transparent',
    backgroundHover: '#f3f2f1',
    color: '#323130',
    borderRadius: '4px',
    padding: '8px 16px',
    fontSize: '14px',
    fontWeight: 600,
  },
} as const;

/**
 * Badge Tokens
 * Status badge styling
 */
export const JMLBadge = {
  borderRadius: '12px',
  padding: '4px 12px',
  fontSize: '12px',
  fontWeight: 600,

  // Variants
  variants: {
    joiner: {
      background: '#e3f2fd',
      color: '#0078d4',
    },
    mover: {
      background: '#f3e5f5',
      color: '#9c27b0',
    },
    leaver: {
      background: '#fff3e0',
      color: '#f57c00',
    },
    completed: {
      background: '#e8f5e9',
      color: '#4caf50',
    },
    inProgress: {
      background: '#fff3e0',
      color: '#f57c00',
    },
    pending: {
      background: '#fff3e0',
      color: '#f57c00',
    },
    overdue: {
      background: '#ffebee',
      color: '#f44336',
    },
    high: {
      background: '#ffebee',
      color: '#f44336',
    },
    medium: {
      background: '#fff3e0',
      color: '#f57c00',
    },
    low: {
      background: '#e8f5e9',
      color: '#4caf50',
    },
  },
} as const;

/**
 * Progress Bar Tokens
 */
export const JMLProgressBar = {
  height: '8px',
  heightLarge: '12px',
  background: '#edebe9',
  fill: 'linear-gradient(90deg, #0078d4 0%, #106ebe 100%)',
  borderRadius: '4px',
  transition: 'width 0.3s ease',
} as const;

/**
 * Grid & Layout Tokens
 */
export const JMLLayout = {
  // Breakpoints
  breakpoints: {
    mobile: '768px',
    tablet: '1024px',
    desktop: '1440px',
    wide: '1920px',
  },

  // Grid gaps
  gap: {
    small: '8px',
    medium: '16px',
    large: '24px',
    xlarge: '32px',
  },

  // Column counts
  columns: {
    auto: 'repeat(auto-fit, minmax(240px, 1fr))',
    two: 'repeat(2, 1fr)',
    three: 'repeat(3, 1fr)',
    four: 'repeat(4, 1fr)',
  },
} as const;

/**
 * Animation Tokens
 */
export const JMLAnimation = {
  // Durations
  duration: {
    fast: '0.1s',
    normal: '0.2s',
    slow: '0.3s',
    verySlow: '0.5s',
  },

  // Easing
  easing: {
    standard: 'ease',
    easeIn: 'ease-in',
    easeOut: 'ease-out',
    easeInOut: 'ease-in-out',
  },
} as const;

/**
 * Z-Index Tokens
 * Layering system
 */
export const JMLZIndex = {
  base: 0,
  dropdown: 1000,
  sticky: 1020,
  fixed: 1030,
  modalBackdrop: 1040,
  modal: 1050,
  popover: 1060,
  tooltip: 1070,
} as const;

/**
 * Utility: Generate CSS custom properties from tokens
 * Use this to inject design tokens into your component styles
 */
export const generateCSSVariables = (): string => {
  return `
    /* JML Design Tokens - CSS Variables */
    :root {
      /* Colors */
      --jml-color-primary: ${JMLColors.themePrimary};
      --jml-color-primary-dark: ${JMLColors.themeDark};
      --jml-color-success: ${JMLColors.success};
      --jml-color-warning: ${JMLColors.warning};
      --jml-color-error: ${JMLColors.error};

      /* Typography */
      --jml-font-family: ${JMLTypography.fontFamily};
      --jml-font-size-title: ${JMLTypography.fontSize.title};
      --jml-font-size-base: ${JMLTypography.fontSize.base};

      /* Spacing */
      --jml-spacing-xs: ${JMLSpacing.xs};
      --jml-spacing-s: ${JMLSpacing.s};
      --jml-spacing-m: ${JMLSpacing.m};
      --jml-spacing-l: ${JMLSpacing.l};
      --jml-spacing-xl: ${JMLSpacing.xl};

      /* Borders */
      --jml-border-radius-medium: ${JMLBorderRadius.medium};
      --jml-border-radius-large: ${JMLBorderRadius.large};

      /* Shadows */
      --jml-shadow-4: ${JMLShadows.depth4};
      --jml-shadow-8: ${JMLShadows.depth8};

      /* Header */
      --jml-header-icon-size: ${JMLHeader.iconSize};
      --jml-header-title-size: ${JMLHeader.titleSize};

      /* Cards */
      --jml-card-border-radius: ${JMLCard.borderRadius};
      --jml-card-shadow: ${JMLCard.shadow};
      --jml-card-shadow-hover: ${JMLCard.shadowHover};
    }
  `;
};

/**
 * Export all tokens as a single object
 */
export const JMLDesignTokens = {
  colors: JMLColors,
  typography: JMLTypography,
  spacing: JMLSpacing,
  borderRadius: JMLBorderRadius,
  shadows: JMLShadows,
  header: JMLHeader,
  card: JMLCard,
  container: JMLContainer,
  button: JMLButton,
  badge: JMLBadge,
  progressBar: JMLProgressBar,
  layout: JMLLayout,
  animation: JMLAnimation,
  zIndex: JMLZIndex,
} as const;

export default JMLDesignTokens;
