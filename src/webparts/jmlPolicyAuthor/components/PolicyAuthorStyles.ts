/**
 * Shared style constants for PolicyAuthorEnhanced and related components.
 * Extracted from inline styles to prevent object recreation on every render.
 */
import * as React from 'react';

// ============================================================================
// COLOR TOKENS
// ============================================================================

export const Colors = {
  /** Secondary text color (descriptions, metadata) */
  textSecondary: '#605e5c',
  /** Tertiary/muted text */
  textMuted: '#6b7280',
  /** Dark heading text */
  textDark: '#1f2937',
  /** Light disabled text */
  textDisabled: '#a19f9d',
  /** Forest Teal primary */
  tealPrimary: '#0d9488',
  /** Forest Teal dark */
  tealDark: '#0f766e',
  /** Microsoft Blue (legacy icons) */
  bluePrimary: '#0078d4',
  /** Success green */
  green: '#059669',
  /** Dark green */
  greenDark: '#064e3b',
  /** Warning amber */
  amber: '#d97706',
  /** Error red */
  red: '#dc2626',
} as const;

// ============================================================================
// TYPOGRAPHY STYLES (reusable objects)
// ============================================================================

export const TextStyles = {
  /** fontWeight: 600 — most common inline style (40+ occurrences) */
  semiBold: { fontWeight: 600 } as React.CSSProperties,
  /** fontWeight: 700 */
  bold: { fontWeight: 700 } as React.CSSProperties,
  /** fontWeight: 500 */
  medium: { fontWeight: 500 } as React.CSSProperties,
  /** Secondary text: color #605e5c */
  secondary: { color: Colors.textSecondary } as React.CSSProperties,
  /** Secondary text with standard body size */
  secondaryBody: { color: Colors.textSecondary, fontSize: 12 } as React.CSSProperties,
  /** Small secondary text */
  secondarySmall: { color: Colors.textSecondary, fontSize: 11 } as React.CSSProperties,
  /** Muted small text */
  mutedSmall: { color: Colors.textMuted, fontSize: 12 } as React.CSSProperties,
  /** Block label: semi-bold, block display, bottom margin */
  blockLabel: { fontWeight: 600, display: 'block' as const, marginBottom: 4, fontSize: 12 } as React.CSSProperties,
  /** Section heading */
  sectionHeading: { fontWeight: 700, color: Colors.textDark, marginBottom: 12, display: 'flex' as const, alignItems: 'center', gap: 6 } as React.CSSProperties,
} as const;

// ============================================================================
// ICON STYLES
// ============================================================================

export const IconStyles = {
  /** Small icon: 12px, muted color */
  small: { fontSize: 12, color: Colors.textMuted } as React.CSSProperties,
  /** Medium icon: 18px, teal */
  mediumTeal: { fontSize: 18, color: Colors.tealPrimary } as React.CSSProperties,
  /** Large icon: 32px, blue */
  largeBlue: { fontSize: 32, color: Colors.bluePrimary, marginBottom: 12 } as React.CSSProperties,
  /** Large icon: 40px, teal */
  largeTeal: { fontSize: 40, color: Colors.tealPrimary, marginBottom: 12 } as React.CSSProperties,
} as const;

// ============================================================================
// LAYOUT STYLES
// ============================================================================

export const LayoutStyles = {
  /** Flex center (horizontal + vertical) */
  flexCenter: { display: 'flex' as const, alignItems: 'center', justifyContent: 'center' } as React.CSSProperties,
  /** Flex row with space-between */
  flexBetween: { display: 'flex' as const, justifyContent: 'space-between', alignItems: 'center' } as React.CSSProperties,
  /** Flex row with center alignment */
  flexRow: { display: 'flex' as const, alignItems: 'center' } as React.CSSProperties,
  /** Standard top padding */
  paddingTop16: { paddingTop: 16 } as React.CSSProperties,
  /** Standard vertical padding */
  paddingVertical16: { padding: '16px 0' } as React.CSSProperties,
  /** Standard bottom margin */
  marginBottom16: { marginBottom: 16 } as React.CSSProperties,
  /** Small bottom margin */
  marginBottom8: { marginBottom: 8 } as React.CSSProperties,
} as const;

// ============================================================================
// BADGE / TAG STYLES
// ============================================================================

export const BadgeStyles = {
  /** Small rounded badge */
  small: { fontSize: 11, fontWeight: 600, padding: '1px 8px', borderRadius: 10 } as React.CSSProperties,
  /** Standard tag */
  tag: { fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 4 } as React.CSSProperties,
} as const;

// ============================================================================
// CONTAINER STYLES
// ============================================================================

export const ContainerStyles = {
  /** Small icon container (28x28) */
  iconSmall: { width: 28, height: 28, display: 'flex' as const, alignItems: 'center', justifyContent: 'center' } as React.CSSProperties,
  /** Medium icon container (32x32) */
  iconMedium: { width: 32, height: 32, display: 'flex' as const, alignItems: 'center', justifyContent: 'center' } as React.CSSProperties,
  /** Teal left border accent */
  tealBorderLeft: { borderLeft: '4px solid #0d9488' } as React.CSSProperties,
} as const;
