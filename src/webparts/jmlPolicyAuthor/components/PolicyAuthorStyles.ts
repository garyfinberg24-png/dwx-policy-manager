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
  /** Darker heading text */
  textDarker: '#111827',
  /** Light disabled text */
  textDisabled: '#a19f9d',
  /** Subtle gray text */
  textSubtle: '#9ca3af',
  /** Forest Teal primary */
  tealPrimary: '#0d9488',
  /** Forest Teal dark */
  tealDark: '#0f766e',
  /** Light teal background */
  tealLight: '#f0fdfa',
  /** Teal border */
  tealBorder: '#99f6e4',
  /** Active teal badge bg */
  tealBadgeBg: '#ccfbf1',
  /** Microsoft Blue (legacy icons) */
  bluePrimary: '#0078d4',
  /** Success green */
  green: '#059669',
  /** Dark green */
  greenDark: '#064e3b',
  /** Light green bg */
  greenLightBg: '#ecfdf5',
  /** Warning amber */
  amber: '#d97706',
  /** Amber light bg */
  amberLightBg: '#fffbeb',
  /** Error red */
  red: '#dc2626',
  /** Required asterisk red */
  redRequired: '#d13438',
  /** Red light bg */
  redLightBg: '#fef2f2',
  /** Border color */
  border: '#e2e8f0',
  /** Light border */
  borderLight: '#edebe9',
  /** Surface gray */
  surfaceGray: '#f3f2f1',
  /** Surface light */
  surfaceLight: '#f8fafc',
} as const;

// ============================================================================
// TYPOGRAPHY STYLES (reusable objects)
// ============================================================================

export const TextStyles = {
  /** fontWeight: 600 — most common inline style */
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
  /** Subtle small text */
  subtleSmall: { fontSize: 10, color: Colors.textSubtle } as React.CSSProperties,
  /** Block label: semi-bold, block display, bottom margin */
  blockLabel: { fontWeight: 600, display: 'block' as const, marginBottom: 4, fontSize: 12 } as React.CSSProperties,
  /** Section heading with flex layout */
  sectionHeading: { fontWeight: 700, color: Colors.textDark, marginBottom: 12, display: 'flex' as const, alignItems: 'center', gap: 6 } as React.CSSProperties,
  /** Section label (semi-bold, flex) */
  sectionLabel: { fontWeight: 600, color: '#334155', flex: 1 } as React.CSSProperties,
  /** Label with icon offset */
  labelWithIconOffset: { marginLeft: 26, color: Colors.textSecondary } as React.CSSProperties,
  /** Bold dark block text */
  boldDarkBlock: { fontWeight: 700, color: Colors.textDarker, display: 'block' as const } as React.CSSProperties,
  /** Muted small with top margin */
  mutedSmallTop: { color: Colors.textMuted, marginTop: 2 } as React.CSSProperties,
} as const;

// ============================================================================
// ICON STYLES
// ============================================================================

export const IconStyles = {
  /** Tiny icon: 10px */
  tiny: { fontSize: 10 } as React.CSSProperties,
  /** Small icon: 11px */
  xSmall: { fontSize: 11 } as React.CSSProperties,
  /** Small icon: 12px, muted color */
  small: { fontSize: 12, color: Colors.textMuted } as React.CSSProperties,
  /** Small icon: 12px, no color */
  small12: { fontSize: 12 } as React.CSSProperties,
  /** Medium icon: 18px, teal */
  mediumTeal: { fontSize: 18, color: Colors.tealPrimary } as React.CSSProperties,
  /** Large icon: 32px, blue */
  largeBlue: { fontSize: 32, color: Colors.bluePrimary, marginBottom: 12 } as React.CSSProperties,
  /** Large icon: 40px, teal */
  largeTeal: { fontSize: 40, color: Colors.tealPrimary, marginBottom: 12 } as React.CSSProperties,
  /** Required field asterisk */
  requiredAsterisk: { fontSize: 8, color: Colors.redRequired } as React.CSSProperties,
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
  /** Flex row with gap 6 */
  flexRowGap6: { display: 'flex' as const, alignItems: 'center', gap: 6 } as React.CSSProperties,
  /** Flex row with gap 8 */
  flexRowGap8: { display: 'flex' as const, alignItems: 'center', gap: 8 } as React.CSSProperties,
  /** Standard top padding */
  paddingTop16: { paddingTop: 16 } as React.CSSProperties,
  /** Standard vertical padding */
  paddingVertical16: { padding: '16px 0' } as React.CSSProperties,
  /** Standard bottom margin */
  marginBottom16: { marginBottom: 16 } as React.CSSProperties,
  /** Small bottom margin */
  marginBottom8: { marginBottom: 8 } as React.CSSProperties,
  /** Small top margin */
  marginTop8: { marginTop: 8 } as React.CSSProperties,
} as const;

// ============================================================================
// BADGE / TAG STYLES
// ============================================================================

export const BadgeStyles = {
  /** Small rounded badge (pills) */
  small: { fontSize: 11, fontWeight: 600, padding: '1px 8px', borderRadius: 10 } as React.CSSProperties,
  /** Standard tag */
  tag: { fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 4 } as React.CSSProperties,
  /** Tiny tag (quiz missing, version badges) */
  tiny: { fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 4 } as React.CSSProperties,
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
  /** Image thumbnail container */
  imageThumbnail: { width: 80, height: 60, borderRadius: 4, overflow: 'hidden' as const, border: '1px solid #e2e8f0' } as React.CSSProperties,
  /** File chip */
  fileChip: { background: Colors.surfaceGray, padding: '6px 12px', borderRadius: 4, display: 'flex' as const, alignItems: 'center' } as React.CSSProperties,
  /** Info/preview box */
  infoBox: { padding: 16, background: Colors.surfaceGray, borderRadius: 8 } as React.CSSProperties,
  /** Content preview with scroll */
  contentPreview: { maxHeight: 300, overflow: 'auto' as const } as React.CSSProperties,
  /** Bordered card */
  borderedCard: { padding: 16, border: '1px solid #e2e8f0', borderRadius: 8 } as React.CSSProperties,
  /** Light bordered preview */
  previewBox: { padding: 12, background: Colors.surfaceLight, borderRadius: 6, border: '1px solid #e2e8f0' } as React.CSSProperties,
} as const;
