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
  greenDark: '#064e3b',
  slateLight: '#94a3b8',
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
} as const;

// ============================================================================
// ICON STYLES
// ============================================================================

export const IconStyles = {
  mediumTeal: { fontSize: 18, color: Colors.tealPrimary } as React.CSSProperties,
  boldTeal: { fontWeight: 700, color: Colors.tealPrimary, display: 'block' as const } as React.CSSProperties,
} as const;

// ============================================================================
// LAYOUT STYLES
// ============================================================================

export const LayoutStyles = {
  paddingTop16: { paddingTop: 16 } as React.CSSProperties,
  paddingVertical16: { padding: '16px 0' } as React.CSSProperties,
  marginBottom16: { marginBottom: 16 } as React.CSSProperties,
  marginTop8: { marginTop: 8 } as React.CSSProperties,
  flex1: { flex: 1 } as React.CSSProperties,
  flex1Center: { flex: 1, textAlign: 'center' as const } as React.CSSProperties,
  textCenter: { textAlign: 'center' as const } as React.CSSProperties,
} as const;

// ============================================================================
// CONTAINER STYLES
// ============================================================================

export const ContainerStyles = {
  tealBorderLeft: { borderLeft: '4px solid #0d9488' } as React.CSSProperties,
  slateLabel: { color: Colors.slateLight, fontWeight: 500 } as React.CSSProperties,
} as const;
