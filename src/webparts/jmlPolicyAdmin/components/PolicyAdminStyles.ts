/**
 * Shared style constants for PolicyAdmin and related components.
 * Extracted from inline styles to prevent object recreation on every render.
 */
import * as React from 'react';

// ============================================================================
// COLOR TOKENS
// ============================================================================

export const Colors = {
  textSecondary: '#605e5c',
  textTertiary: '#64748b',
  textDark: '#0f172a',
  textSlate: '#475569',
  tealPrimary: '#0d9488',
  tealLight: '#f0fdfa',
  tealBorder: '#99f6e4',
  tealBadgeBg: '#ccfbf1',
  greenDark: '#064e3b',
  slateLight: '#94a3b8',
  /** Surface gray background */
  surfaceGray: '#f1f5f9',
  /** Light surface */
  surfaceLight: '#f8fafc',
  /** Card border */
  border: '#e2e8f0',
  /** Light border */
  borderLight: '#edebe9',
  /** Warning amber */
  amber: '#d97706',
  /** Error red */
  red: '#dc2626',
  /** Success green */
  green: '#10b981',
  /** Purple accent */
  purple: '#7c3aed',
  /** Purple light bg */
  purpleLightBg: '#ede9fe',
  /** Indigo accent */
  indigo: '#6366f1',
} as const;

// ============================================================================
// TYPOGRAPHY STYLES
// ============================================================================

export const TextStyles = {
  semiBold: { fontWeight: 600 } as React.CSSProperties,
  bold: { fontWeight: 700 } as React.CSSProperties,
  medium: { fontWeight: 500 } as React.CSSProperties,
  secondary: { color: Colors.textSecondary } as React.CSSProperties,
  tertiary: { color: Colors.textTertiary } as React.CSSProperties,
  blockLabel: { fontWeight: 600, display: 'block' as const, marginBottom: 12 } as React.CSSProperties,
  primaryDark: { fontWeight: 500, color: Colors.textDark } as React.CSSProperties,
  /** Section header: semi-bold block with margin */
  sectionHeader: { fontWeight: 600, marginBottom: 12, display: 'block' as const } as React.CSSProperties,
  /** Clickable text link */
  clickableText: { fontWeight: 500, color: Colors.textDark, cursor: 'pointer', textDecoration: 'underline' } as React.CSSProperties,
  /** Monospace muted text (code/subjects) */
  monospaceMuted: { fontFamily: 'monospace', fontSize: 11, color: Colors.textTertiary } as React.CSSProperties,
  /** Small muted text */
  smallMuted: { fontSize: 11, color: Colors.textTertiary } as React.CSSProperties,
  /** Small secondary text */
  smallSecondary: { fontSize: 12, color: Colors.textSecondary } as React.CSSProperties,
  /** Slate label text */
  slateLabel: { color: Colors.slateLight, fontWeight: 500 } as React.CSSProperties,
  /** Small slate text */
  smallSlate: { fontSize: 11, color: Colors.slateLight } as React.CSSProperties,
} as const;

// ============================================================================
// ICON STYLES
// ============================================================================

export const IconStyles = {
  /** Small icon: 12px */
  small: { fontSize: 12 } as React.CSSProperties,
  /** Small-medium icon: 14px */
  smallMedium: { fontSize: 14 } as React.CSSProperties,
  /** Medium icon: 16px */
  medium: { fontSize: 16 } as React.CSSProperties,
  /** Medium-large icon: 18px */
  mediumLarge: { fontSize: 18 } as React.CSSProperties,
  /** Medium icon teal */
  mediumTeal: { fontSize: 18, color: Colors.tealPrimary } as React.CSSProperties,
  /** Large icon: 20px */
  large: { fontSize: 20 } as React.CSSProperties,
  /** XLarge icon: 22px */
  xLarge: { fontSize: 22 } as React.CSSProperties,
  /** XXLarge icon: 24px */
  xxLarge: { fontSize: 24 } as React.CSSProperties,
  /** Jumbo icon: 48px */
  jumbo: { fontSize: 48 } as React.CSSProperties,
  /** Bold teal */
  boldTeal: { fontWeight: 700, color: Colors.tealPrimary, display: 'block' as const } as React.CSSProperties,
} as const;

// ============================================================================
// LAYOUT STYLES
// ============================================================================

export const LayoutStyles = {
  paddingTop16: { paddingTop: 16 } as React.CSSProperties,
  paddingTop12: { paddingTop: 12 } as React.CSSProperties,
  paddingVertical16: { padding: '16px 0' } as React.CSSProperties,
  marginBottom16: { marginBottom: 16 } as React.CSSProperties,
  marginBottom8: { marginBottom: 8 } as React.CSSProperties,
  marginTop8: { marginTop: 8 } as React.CSSProperties,
  flex1: { flex: 1 } as React.CSSProperties,
  flex1Center: { flex: 1, textAlign: 'center' as const } as React.CSSProperties,
  textCenter: { textAlign: 'center' as const } as React.CSSProperties,
  /** Flex row with center alignment */
  flexRow: { display: 'flex' as const, alignItems: 'center' } as React.CSSProperties,
  /** Flex row with gap 8 */
  flexRowGap8: { display: 'flex' as const, alignItems: 'center', gap: 8 } as React.CSSProperties,
  /** Flex row with gap 12 */
  flexRowGap12: { display: 'flex' as const, alignItems: 'center', gap: 12 } as React.CSSProperties,
  /** Flex row with gap 12 and bottom margin */
  flexRowGap12Mb8: { display: 'flex' as const, alignItems: 'center', gap: 12, marginBottom: 8 } as React.CSSProperties,
  /** Flex center (horizontal + vertical) */
  flexCenter: { display: 'flex' as const, alignItems: 'center', justifyContent: 'center' } as React.CSSProperties,
} as const;

// ============================================================================
// BADGE / STATUS STYLES
// ============================================================================

export const BadgeStyles = {
  /** Active/inactive toggle badge */
  activeInactive: { padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600 } as React.CSSProperties,
  /** Small rounded pill */
  pill: { padding: '2px 8px', borderRadius: 12, fontSize: 11, fontWeight: 600 } as React.CSSProperties,
  /** Standard tag */
  tag: { padding: '2px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600 } as React.CSSProperties,
  /** Department chip */
  departmentChip: { padding: '1px 6px', borderRadius: 3, fontSize: 10, fontWeight: 500, background: Colors.tealLight, color: Colors.tealPrimary, border: '1px solid ' + Colors.tealBorder } as React.CSSProperties,
  /** Highlight badge (bold, rounded) */
  highlight: { padding: '3px 10px', borderRadius: 4, fontSize: 11, fontWeight: 700 } as React.CSSProperties,
  /** Default purple pill */
  defaultPurple: { padding: '2px 8px', borderRadius: 12, fontSize: 11, fontWeight: 600, backgroundColor: Colors.purpleLightBg, color: Colors.purple } as React.CSSProperties,
} as const;

// ============================================================================
// CONTAINER STYLES
// ============================================================================

export const ContainerStyles = {
  /** Teal left border accent */
  tealBorderLeft: { borderLeft: '4px solid #0d9488' } as React.CSSProperties,
  /** Color swatch (16x16) */
  colorSwatch: { width: 16, height: 16, borderRadius: 4 } as React.CSSProperties,
  /** Large color swatch (24x24) */
  colorSwatchLarge: { width: 24, height: 24, borderRadius: 4, border: '1px solid #e2e8f0' } as React.CSSProperties,
  /** Preview box (small) */
  previewBox: { padding: 12, background: Colors.surfaceLight, borderRadius: 4, border: '1px solid #e2e8f0' } as React.CSSProperties,
  /** Preview box (large) */
  previewBoxLarge: { padding: 16, background: Colors.surfaceLight, borderRadius: 8, border: '1px solid #e2e8f0' } as React.CSSProperties,
  /** Light teal background */
  tealLightBg: { background: Colors.tealLight } as React.CSSProperties,
  /** Icon container (40x40) */
  iconContainer40: { width: 40, height: 40, borderRadius: 8, background: Colors.tealLight, display: 'flex' as const, alignItems: 'center', justifyContent: 'center' } as React.CSSProperties,
} as const;

// ============================================================================
// KPI / STAT STYLES
// ============================================================================

export const KPIStyles = {
  /** Stat card container */
  statCard: { padding: 20, borderRadius: 8, background: Colors.surfaceLight, display: 'flex' as const, flexDirection: 'column' as const, alignItems: 'center', justifyContent: 'center' } as React.CSSProperties,
} as const;

// ============================================================================
// CARD BORDER STYLES
// ============================================================================

export const CardBorderStyles = {
  /** Warning left border */
  warningLeft: { borderLeft: '4px solid #d97706' } as React.CSSProperties,
  /** AI/indigo left border */
  indigoLeft: { borderLeft: '4px solid #6366f1' } as React.CSSProperties,
} as const;

// ============================================================================
// DIVIDER STYLES
// ============================================================================

export const DividerStyles = {
  /** Vertical line separator */
  verticalLine: { width: 1, height: 40, background: Colors.border } as React.CSSProperties,
  /** Progress bar container */
  progressContainer: { width: '100%', height: 6, borderRadius: 3, background: Colors.border, overflow: 'hidden' as const } as React.CSSProperties,
  /** Section divider */
  sectionDivider: { textAlign: 'center' as const, padding: '16px 0', borderTop: '1px solid #e2e8f0' } as React.CSSProperties,
} as const;

// ============================================================================
// EMAIL TEMPLATE STYLES
// ============================================================================

export const EmailTemplateStyles = {
  /** Template name (clickable) */
  templateName: { fontWeight: 600, fontSize: 13, color: Colors.textDark, cursor: 'pointer' } as React.CSSProperties,
  /** Subject line (monospace) */
  subjectMono: { fontFamily: 'monospace', fontSize: 11, color: Colors.textTertiary } as React.CSSProperties,
  /** Merge tag pill */
  mergeTagPill: { padding: '4px 8px', borderRadius: 4, backgroundColor: Colors.purpleLightBg, color: Colors.purple, fontFamily: 'monospace', fontSize: 11, fontWeight: 500 } as React.CSSProperties,
} as const;
