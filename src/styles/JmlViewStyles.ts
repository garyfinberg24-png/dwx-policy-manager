// @ts-nocheck
/* eslint-disable */
/**
 * JML View Layout Styles - COMPLETE Global Style Definitions
 *
 * ╔══════════════════════════════════════════════════════════════════════════════════╗
 * ║ AUTHORITATIVE SOURCE: docs/mockups/JML-Global-Layout-Standards-Preview.html      ║
 * ║ All values in this file MUST match the official JML Styling Reference Guide      ║
 * ║ EVERY view component MUST import from this file for consistency                  ║
 * ╚══════════════════════════════════════════════════════════════════════════════════╝
 *
 * COVERED ELEMENTS:
 * 1. Page Header (gradient blue, breadcrumb, title)
 * 2. Subheader Panels (3 styles: accent, underline, banner)
 * 3. Main Content Container
 * 4. Stats/KPI Grid and Cards
 * 5. Table Styles (container, header, rows, cells, status badges)
 * 6. Command Panel / Action Bar
 * 7. Cards (standard, interactive, KPI)
 * 8. Empty & Loading States
 * 9. Filter Row
 * 10. Section Headers
 * 11. Grid Layouts
 *
 * Tab Panel styles are in TabPanelStyles.ts (useTabPanelStyles hook)
 *
 * USAGE:
 * import { JmlPageHeaderStyles, JmlSubheaderStyles, JmlTableStyles, ... } from '../../../../styles';
 */

import { FluentColors, FluentSpacing, FluentTypography, FluentShadows, FluentBorderRadius, FluentAnimations } from './FluentUIStyles';

// ════════════════════════════════════════════════════════════════════════════════════
// 1. PAGE HEADER STYLES
// Source: .jml-page-header (lines 312-367)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlPageHeaderStyles = {
  /**
   * Page header container - gradient blue background
   * Source: .jml-page-header { padding: 24px 40px 24px 72px; background: linear-gradient... }
   */
  container: {
    padding: '24px 40px 24px 72px',
    background: 'linear-gradient(180deg, #004578 0%, #005a9e 100%)',
    borderTop: '1px solid rgba(255, 255, 255, 0.2)',
    borderBottom: `4px solid ${FluentColors.themePrimary}`,
    minHeight: '100px',
    color: FluentColors.white,
  },

  /**
   * Breadcrumb navigation
   * Source: .jml-breadcrumb { display: flex; gap: 8px; font-size: 12px; }
   */
  breadcrumb: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '12px',
    marginBottom: '8px',
  },

  breadcrumbLink: {
    color: 'rgba(255, 255, 255, 0.8)',
    textDecoration: 'none',
    ':hover': {
      color: FluentColors.white,
      textDecoration: 'underline',
    },
  },

  breadcrumbSeparator: {
    color: 'rgba(255, 255, 255, 0.5)',
  },

  breadcrumbCurrent: {
    color: FluentColors.white,
  },

  /**
   * Page title
   * Source: .jml-page-title { font-size: 28px; font-weight: 600; }
   */
  title: {
    fontSize: '28px',
    fontWeight: 600,
    margin: '0 0 8px 0',
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    textShadow: '0 1px 2px rgba(0, 0, 0, 0.2)',
    fontFamily: FluentTypography.fontFamily,
  },

  titleIcon: {
    fontSize: '32px',
  },

  /**
   * Page description
   * Source: .jml-page-description { font-size: 14px; color: rgba(255,255,255,0.9); }
   */
  description: {
    fontSize: '14px',
    color: 'rgba(255, 255, 255, 0.9)',
    margin: 0,
    maxWidth: '600px',
    lineHeight: 1.5,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 2. SUBHEADER PANEL STYLES (3 variants)
// Source: .jml-subheader-* (lines 432-519)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlSubheaderStyles = {
  /**
   * Style 1: Blue Left Accent (Default)
   * Source: .jml-subheader-accent { background: linear-gradient; border-left: 4px solid; }
   */
  accent: {
    background: 'linear-gradient(135deg, #e8f4fd 0%, #f0f8ff 100%)',
    borderLeft: `4px solid ${FluentColors.themePrimary}`,
    borderRadius: '8px',
    padding: '16px 24px',
  },

  accentTitle: {
    fontSize: '20px',
    fontWeight: 600,
    color: '#004578',
    margin: '0 0 4px 0',
    fontFamily: FluentTypography.fontFamily,
  },

  accentSubtitle: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  /**
   * Style 2: Underline with Step Badge (for wizards)
   * Source: .jml-subheader-underline { border-bottom: 3px solid; display: flex; }
   */
  underline: {
    backgroundColor: FluentColors.white,
    borderBottom: `3px solid ${FluentColors.themePrimary}`,
    padding: '8px 16px',
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },

  stepBadge: {
    backgroundColor: FluentColors.themePrimary,
    color: FluentColors.white,
    width: '28px',
    height: '28px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '14px',
    fontWeight: 600,
    flexShrink: 0,
  },

  underlineTitle: {
    fontSize: '18px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  /**
   * Style 3: Icon Banner (Modern)
   * Source: .jml-subheader-banner { background: white; border: 1px solid; border-radius: 12px; }
   */
  banner: {
    backgroundColor: FluentColors.white,
    border: '1px solid #e1e5e9',
    borderRadius: '12px',
    padding: '24px 32px',
    boxShadow: '0 2px 8px rgba(0, 0, 0, 0.06)',
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
  },

  bannerIconBox: {
    background: 'linear-gradient(135deg, #0078d4 0%, #004578 100%)',
    width: '52px',
    height: '52px',
    borderRadius: '12px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: FluentColors.white,
    fontSize: '24px',
    flexShrink: 0,
  },

  bannerContent: {
    flex: 1,
  },

  bannerTitle: {
    fontSize: '20px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: '0 0 4px 0',
    fontFamily: FluentTypography.fontFamily,
  },

  bannerSubtitle: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 3. MAIN CONTENT CONTAINER
// Source: .jml-main-content (lines 527-529)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlViewStyles = {
  /**
   * Standard view container with horizontal padding
   * Source: .jml-main-content { background: white; padding: 24px; }
   */
  viewContainer: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '20px',
    padding: '0 24px',
  },

  viewContainerNoPadding: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '20px',
    padding: 0,
  },

  viewContainerFullWidth: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '20px',
    width: '100%',
    maxWidth: '100%',
    padding: 0,
  },

  mainContent: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '16px',
    flex: 1,
  },

  section: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '12px',
  },

  sectionSpaced: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '12px',
    marginTop: '24px',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 4. STATS / KPI GRID AND CARDS
// Source: .card-kpi, .kpi-value, .kpi-label, .kpi-trend (lines 1047-1052)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlStatsRowStyles = {
  /**
   * Stats row container - 4-column grid for KPI cards
   * Used by: ApprovalsView, PartiesView, ObligationsView, etc.
   */
  container: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '16px',
    '@media (max-width: 1200px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 600px)': {
      gridTemplateColumns: '1fr',
    },
  },

  /**
   * 6-column responsive grid
   */
  statsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(6, 1fr)',
    gap: '12px',
    // NOTE: paddingTop removed - let component control spacing via margin
    '@media (max-width: 1400px)': {
      gridTemplateColumns: 'repeat(4, 1fr)',
    },
    '@media (max-width: 1000px)': {
      gridTemplateColumns: 'repeat(3, 1fr)',
    },
    '@media (max-width: 700px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 480px)': {
      gridTemplateColumns: '1fr',
    },
  },

  statsGrid4: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '12px',
    // NOTE: paddingTop removed - let component control spacing via margin
    '@media (max-width: 1000px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 480px)': {
      gridTemplateColumns: '1fr',
    },
  },

  statsGrid3: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '12px',
    // NOTE: paddingTop removed - let component control spacing via margin
    '@media (max-width: 700px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 480px)': {
      gridTemplateColumns: '1fr',
    },
  },

  /**
   * Standard stat card container
   * Source: .card-kpi { text-align: center; padding: 24px; }
   */
  statCard: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    padding: '24px',
    boxShadow: FluentShadows.depth4,
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    minHeight: '100px',
    textAlign: 'center' as const,
    transition: `all ${FluentAnimations.durationNormal} ${FluentAnimations.easeInOut}`,
    cursor: 'pointer',
    ':hover': {
      boxShadow: FluentShadows.depth8,
      transform: 'translateY(-2px)',
      borderColor: FluentColors.themePrimary,
    },
  },

  statCardStatic: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    padding: '24px',
    boxShadow: FluentShadows.depth4,
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    minHeight: '100px',
    textAlign: 'center' as const,
  },

  /**
   * Large stat value
   * Source: .kpi-value { font-size: 36px; font-weight: 700; line-height: 1; margin-bottom: 8px; }
   */
  statValue: {
    fontSize: '36px',
    fontWeight: 700,
    lineHeight: 1,
    marginBottom: '8px',
    fontVariantNumeric: 'tabular-nums',
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },

  /**
   * Stat label
   * Source: .kpi-label { font-size: 14px; color: var(--jml-neutral-secondary); }
   */
  statLabel: {
    fontSize: '14px',
    fontWeight: 500,
    color: FluentColors.neutralSecondary,
    fontFamily: FluentTypography.fontFamily,
    lineHeight: 1.3,
  },

  /**
   * Stat trend indicator
   * Source: .kpi-trend { font-size: 12px; margin-top: 8px; }
   */
  statTrend: {
    fontSize: '12px',
    fontWeight: 500,
    marginTop: '8px',
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },

  statTrendPositive: {
    color: FluentColors.success,
  },

  statTrendNegative: {
    color: FluentColors.error,
  },

  statTrendNeutral: {
    color: FluentColors.neutralTertiary,
  },

  statIcon: {
    fontSize: '20px',
    color: FluentColors.themePrimary,
    marginBottom: '8px',
  },

  /**
   * Stat content wrapper - holds value and label
   * Used by: ApprovalsView, PartiesView, ObligationsView, etc.
   */
  statContent: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
    flex: 1,
    minWidth: 0,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 5. TABLE STYLES
// Source: .jml-table-* (lines 672-730, 1377-1532)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlTableStyles = {
  /**
   * Table container with border and rounded corners
   * Source: .jml-table-container { border: 1px solid; border-radius: 8px; }
   */
  container: {
    border: `1px solid ${FluentColors.neutralLight}`,
    borderRadius: '8px',
    overflow: 'hidden',
    backgroundColor: FluentColors.white,
  },

  /**
   * Table card wrapper with header
   * Source: .jml-table-card { background: white; border-radius: 12px; box-shadow; }
   */
  card: {
    backgroundColor: FluentColors.white,
    borderRadius: '12px',
    boxShadow: '0 2px 8px rgba(0, 0, 0, 0.08)',
    overflow: 'hidden',
  },

  /**
   * Table card header
   * Source: .jml-table-header { display: flex; align-items: center; padding: 20px 24px; }
   */
  cardHeader: {
    display: 'flex',
    alignItems: 'center',
    padding: '20px 24px',
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
    gap: '16px',
  },

  cardHeaderIcon: {
    width: '40px',
    height: '40px',
    borderRadius: '10px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: FluentColors.white,
    fontSize: '20px',
    flexShrink: 0,
  },

  cardHeaderIconBlue: {
    backgroundColor: FluentColors.themePrimary,
  },

  cardHeaderIconTeal: {
    backgroundColor: '#03787C',
  },

  cardHeaderIconGreen: {
    backgroundColor: FluentColors.success,
  },

  cardHeaderIconOrange: {
    backgroundColor: '#f57c00',
  },

  cardHeaderContent: {
    flex: 1,
  },

  cardHeaderTitle: {
    fontSize: '16px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  cardHeaderSubtitle: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    margin: '4px 0 0 0',
    fontFamily: FluentTypography.fontFamily,
  },

  cardHeaderActions: {
    display: 'flex',
    gap: '8px',
  },

  /**
   * Base table styles
   * Source: .jml-table { width: 100%; border-collapse: collapse; }
   */
  table: {
    width: '100%',
    borderCollapse: 'collapse' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  /**
   * Table header
   * Source: .jml-table thead { background: #f3f2f1; position: relative; }
   */
  thead: {
    backgroundColor: FluentColors.neutralLighter,
    position: 'relative' as const,
  },

  theadWithAccent: {
    backgroundColor: FluentColors.neutralLighter,
    position: 'relative' as const,
    '::before': {
      content: '""',
      position: 'absolute' as const,
      left: 0,
      top: 0,
      bottom: 0,
      width: '4px',
      backgroundColor: FluentColors.themePrimary,
    },
  },

  /**
   * Table header cell
   * Source: .jml-table th { text-align: left; padding: 14px 16px; font-size: 12px; font-weight: 600; }
   */
  th: {
    textAlign: 'left' as const,
    padding: '14px 16px',
    fontSize: '12px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
  },

  thFirst: {
    paddingLeft: '20px',
  },

  /**
   * Table data cell
   * Source: .jml-table td { padding: 12px 16px; border-bottom: 1px solid; font-size: 14px; }
   */
  td: {
    padding: '12px 16px',
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
    color: FluentColors.neutralPrimary,
    fontSize: '14px',
  },

  tdFirst: {
    paddingLeft: '20px',
  },

  /**
   * Table row
   */
  tr: {
    transition: 'background-color 0.15s ease',
    ':hover': {
      backgroundColor: '#eff6fc', // Primary lightest
    },
  },

  trLast: {
    '& td': {
      borderBottom: 'none',
    },
  },

  /**
   * Status badges for tables
   * Source: .jml-table-status { padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; }
   */
  statusBadge: {
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '11px',
    fontWeight: 600,
    display: 'inline-block',
  },

  statusPending: {
    backgroundColor: '#fff4ce',
    color: '#986f0b',
  },

  statusActive: {
    backgroundColor: '#deecf9',
    color: '#0078d4',
  },

  statusCompleted: {
    backgroundColor: '#dff6dd',
    color: '#107c10',
  },

  statusOverdue: {
    backgroundColor: '#fde7e9',
    color: '#d13438',
  },

  /**
   * Table action buttons
   * Source: .jml-table-action-btn { width: 28px; height: 28px; border-radius: 4px; }
   */
  actionButton: {
    width: '28px',
    height: '28px',
    borderRadius: '4px',
    border: 'none',
    backgroundColor: 'transparent',
    color: FluentColors.neutralSecondary,
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'all 0.15s ease',
    ':hover': {
      backgroundColor: FluentColors.neutralLighter,
      color: FluentColors.themePrimary,
    },
  },

  /**
   * Table empty state
   */
  emptyState: {
    padding: '48px 24px',
    textAlign: 'center' as const,
    color: FluentColors.neutralSecondary,
  },

  emptyStateIcon: {
    fontSize: '48px',
    color: FluentColors.neutralTertiaryAlt,
    marginBottom: '16px',
  },

  emptyStateTitle: {
    fontSize: '16px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    marginBottom: '8px',
  },

  emptyStateDescription: {
    fontSize: '14px',
    color: FluentColors.neutralSecondary,
    maxWidth: '300px',
    margin: '0 auto',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 6. COMMAND PANEL / ACTION BAR
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlCommandPanelStyles = {
  /**
   * Command panel container
   */
  container: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px 24px',
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    boxShadow: FluentShadows.depth4,
    marginBottom: '16px',
  },

  /**
   * Left side - filters and search
   */
  filters: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    flex: 1,
  },

  /**
   * Search field
   */
  searchField: {
    flex: '0 1 300px',
    minWidth: '200px',
  },

  /**
   * Filter dropdown
   */
  filterDropdown: {
    flex: '0 0 auto',
    minWidth: '150px',
  },

  /**
   * Right side - action buttons
   */
  actions: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    marginLeft: 'auto',
  },

  /**
   * Primary action button (e.g., "New", "Create")
   */
  primaryButton: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 20px',
    backgroundColor: FluentColors.themePrimary,
    color: FluentColors.white,
    border: 'none',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    ':hover': {
      backgroundColor: FluentColors.themeDarkAlt,
    },
  },

  /**
   * Secondary action button
   */
  secondaryButton: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 20px',
    backgroundColor: 'transparent',
    color: FluentColors.neutralPrimary,
    border: `1px solid ${FluentColors.neutralTertiaryAlt}`,
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 500,
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    ':hover': {
      backgroundColor: FluentColors.neutralLighter,
      borderColor: FluentColors.neutralPrimary,
    },
  },

  /**
   * Icon-only action button
   */
  iconButton: {
    width: '36px',
    height: '36px',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'transparent',
    color: FluentColors.neutralSecondary,
    border: 'none',
    borderRadius: '4px',
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    ':hover': {
      backgroundColor: FluentColors.neutralLighter,
      color: FluentColors.neutralPrimary,
    },
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 7. CARD STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlCardStyles = {
  // =========================================================================
  // STANDARD CARD (CMP-CRD)
  // Basic card with header and body
  // =========================================================================
  card: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    padding: '20px',
    boxShadow: FluentShadows.depth4,
  },

  cardInteractive: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    padding: '20px',
    boxShadow: FluentShadows.depth4,
    cursor: 'pointer',
    transition: `all ${FluentAnimations.durationNormal} ${FluentAnimations.easeInOut}`,
    ':hover': {
      boxShadow: FluentShadows.depth16,
      transform: 'translateY(-2px)',
      borderColor: FluentColors.themePrimary,
    },
  },

  cardHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '12px',
    paddingBottom: '12px',
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
  },

  cardTitle: {
    fontSize: '16px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: 0,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontFamily: FluentTypography.fontFamily,
  },

  cardBody: {
    flex: 1,
  },

  cardFooter: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    marginTop: '12px',
    paddingTop: '12px',
    borderTop: `1px solid ${FluentColors.neutralLight}`,
  },

  // =========================================================================
  // ACCENT CARDS - JML Process Colors
  // Left border accent with process colors
  // =========================================================================
  cardAccent: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    borderLeft: '4px solid #0078d4',
    padding: '20px',
    boxShadow: FluentShadows.depth4,
  },

  /**
   * Joiner accent - Blue (#0078d4)
   */
  cardAccentJoiner: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    borderLeft: '4px solid #0078d4',
    padding: '20px',
    boxShadow: FluentShadows.depth4,
  },

  /**
   * Mover accent - Purple (#5c2d91)
   */
  cardAccentMover: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    borderLeft: '4px solid #5c2d91',
    padding: '20px',
    boxShadow: FluentShadows.depth4,
  },

  /**
   * Leaver accent - Orange (#d83b01)
   */
  cardAccentLeaver: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    borderLeft: '4px solid #d83b01',
    padding: '20px',
    boxShadow: FluentShadows.depth4,
  },

  /**
   * Teal accent - JML Teal (#03787C)
   */
  cardAccentTeal: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    borderLeft: '4px solid #03787C',
    padding: '20px',
    boxShadow: FluentShadows.depth4,
  },

  // =========================================================================
  // FANCY CARD
  // Gradient header for featured/premium content
  // =========================================================================
  cardFancy: {
    backgroundColor: '#ffffff',
    background: 'linear-gradient(135deg, #f8fbff 0%, #ffffff 100%)',
    borderRadius: '8px',
    border: 'none',
    boxShadow: '0 4px 12px rgba(0, 0, 0, 0.08)',
    overflow: 'hidden',
  },

  cardFancyHeader: {
    background: 'linear-gradient(135deg, #004578 0%, #0078d4 100%)',
    color: '#ffffff',
    padding: '16px 20px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },

  cardFancyTitle: {
    fontSize: '16px',
    fontWeight: 600,
    color: '#ffffff',
    margin: 0,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },

  cardFancyBody: {
    padding: '20px',
  },

  // =========================================================================
  // KPI CARD
  // Stat display with value, label, and trend
  // =========================================================================
  cardKpi: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    padding: '24px',
    boxShadow: FluentShadows.depth4,
    textAlign: 'center' as const,
  },

  kpiValue: {
    fontSize: '36px',
    fontWeight: 700,
    color: '#0078d4',
    lineHeight: 1,
    marginBottom: '8px',
    fontFamily: FluentTypography.fontFamily,
  },

  kpiLabel: {
    fontSize: '14px',
    color: '#605e5c',
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  kpiTrend: {
    fontSize: '12px',
    marginTop: '8px',
    fontWeight: 500,
    fontFamily: FluentTypography.fontFamily,
  },

  kpiTrendUp: {
    fontSize: '12px',
    marginTop: '8px',
    fontWeight: 500,
    color: '#107c10',
  },

  kpiTrendDown: {
    fontSize: '12px',
    marginTop: '8px',
    fontWeight: 500,
    color: '#a4262c',
  },

  // =========================================================================
  // DASHBOARD TILE
  // Larger card for dashboard widgets
  // =========================================================================
  dashboardTile: {
    backgroundColor: FluentColors.white,
    borderRadius: '12px',
    border: `1px solid ${FluentColors.neutralLight}`,
    boxShadow: FluentShadows.depth4,
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column' as const,
    height: '100%',
  },

  dashboardTileHeader: {
    padding: '16px 20px',
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    backgroundColor: '#faf9f8',
  },

  dashboardTileTitle: {
    fontSize: '14px',
    fontWeight: 600,
    color: '#323130',
    margin: 0,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },

  dashboardTileBody: {
    padding: '20px',
    flex: 1,
    overflow: 'auto' as const,
  },

  dashboardTileFooter: {
    padding: '12px 20px',
    borderTop: `1px solid ${FluentColors.neutralLight}`,
    backgroundColor: '#faf9f8',
  },

  // =========================================================================
  // CARD GRID LAYOUTS
  // =========================================================================
  cardGrid2: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
  },

  cardGrid3: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '16px',
  },

  cardGrid4: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '16px',
  },

  cardGridAuto: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
    gap: '16px',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 8. EMPTY & LOADING STATES
// Source: .empty-state-* (lines 1270-1280)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlEmptyStateStyles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
    textAlign: 'center' as const,
    minHeight: '200px',
  },

  icon: {
    fontSize: '48px',
    color: FluentColors.neutralTertiary,
    marginBottom: '16px',
  },

  title: {
    fontSize: '18px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: '0 0 8px 0',
    fontFamily: FluentTypography.fontFamily,
  },

  message: {
    fontSize: '14px',
    color: FluentColors.neutralSecondary,
    maxWidth: '400px',
    margin: '0 auto 16px auto',
    fontFamily: FluentTypography.fontFamily,
  },

  action: {
    marginTop: '16px',
  },
};

export const JmlLoadingStyles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
    minHeight: '200px',
  },

  spinner: {
    marginBottom: '16px',
  },

  message: {
    fontSize: '14px',
    color: FluentColors.neutralSecondary,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 9. FILTER ROW
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlFilterStyles = {
  filterRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    flexWrap: 'wrap' as const,
    marginBottom: '16px',
  },

  searchField: {
    flex: '1 1 300px',
    minWidth: '200px',
    maxWidth: '400px',
  },

  filterDropdown: {
    flex: '0 0 auto',
    minWidth: '150px',
  },

  filterActions: {
    marginLeft: 'auto',
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 10. SECTION HEADERS
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlSectionStyles = {
  sectionHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '12px',
    paddingBottom: '8px',
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
  },

  sectionTitle: {
    fontSize: '20px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  sectionSubtitle: {
    fontSize: '14px',
    color: FluentColors.neutralSecondary,
    marginTop: '4px',
    fontFamily: FluentTypography.fontFamily,
  },

  sectionActions: {
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 11. GRID LAYOUTS
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlGridStyles = {
  twoColumn: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr',
    },
  },

  twoColumnWide: {
    display: 'grid',
    gridTemplateColumns: '2fr 1fr',
    gap: '16px',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr',
    },
  },

  threeColumn: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '16px',
    '@media (max-width: 1000px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 600px)': {
      gridTemplateColumns: '1fr',
    },
  },

  autoFitGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))',
    gap: '16px',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// UNIFIED EXPORT
// ════════════════════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════════════════════
// 12. JML COLOR PALETTE
// Source: CSS Custom Properties (lines 13-48)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlColorPalette = {
  // Primary Brand Colors
  primary: '#0078d4',
  primaryDark: '#106ebe',
  primaryDarker: '#005a9e',
  primaryDarkest: '#004578',
  primaryLight: '#c7e0f4',
  primaryLighter: '#deecf9',
  primaryLightest: '#eff6fc',

  // Action Color (Teal) - For CTAs
  actionTeal: '#03787C',
  tealDark: '#026569',
  tealLight: '#e0f5f5',

  // Accent Color (Gold) - For Premium indicators
  accentGold: '#d4a017',
  goldDark: '#b8860b',
  goldLight: '#fff8e1',

  // Semantic Colors
  success: '#107c10',
  warning: '#ffb900',
  error: '#d13438',
  errorDark: '#a52a2d',
  info: '#0078d4',

  // Neutral Palette
  neutralPrimary: '#323130',
  neutralSecondary: '#605e5c',
  neutralTertiary: '#8a8886',
  neutralQuaternary: '#a19f9d',
  borderDefault: '#c8c6c4',
  neutralLight: '#edebe9',
  neutralLighter: '#f3f2f1',
  neutralLightest: '#faf9f8',
  white: '#ffffff',

  // Badge Colors
  badge: {
    pending: { bg: '#fff4ce', text: '#986f0b' },
    active: { bg: '#deecf9', text: '#0078d4' },
    inProgress: { bg: '#deecf9', text: '#0078d4' },
    completed: { bg: '#dff6dd', text: '#107c10' },
    overdue: { bg: '#fde7e9', text: '#d13438' },
    blocked: { bg: '#f3f2f1', text: '#605e5c' },
    cancelled: { bg: '#f3f2f1', text: '#8a8886' },
    // Process types
    joiner: { bg: '#deecf9', text: '#0078d4' },
    mover: { bg: '#f3e5f5', text: '#7b1fa2' },
    leaver: { bg: '#fff3e0', text: '#e65100' },
    // Priority
    priorityHigh: { bg: '#fde7e9', text: '#d13438' },
    priorityMedium: { bg: '#fff4ce', text: '#986f0b' },
    priorityLow: { bg: '#dff6dd', text: '#107c10' },
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 13. TYPOGRAPHY STYLES
// Source: CSS Custom Properties (lines 50-68)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlTypographyStyles = {
  // Font Family
  fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",

  // Font Sizes
  fontSizes: {
    hero: '32px',
    title: '28px',
    h1: '24px',
    h2: '20px',
    h3: '18px',
    h4: '16px',
    body: '14px',
    small: '13px',
    caption: '12px',
    tiny: '11px',
    micro: '10px',
  },

  // Font Weights
  fontWeights: {
    regular: 400,
    medium: 500,
    semibold: 600,
    bold: 700,
  },

  // Pre-composed text styles
  hero: {
    fontSize: '32px',
    fontWeight: 700,
    lineHeight: 1.2,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  title: {
    fontSize: '28px',
    fontWeight: 600,
    lineHeight: 1.3,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  h1: {
    fontSize: '24px',
    fontWeight: 600,
    lineHeight: 1.3,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  h2: {
    fontSize: '20px',
    fontWeight: 600,
    lineHeight: 1.3,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  h3: {
    fontSize: '18px',
    fontWeight: 600,
    lineHeight: 1.4,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  h4: {
    fontSize: '16px',
    fontWeight: 600,
    lineHeight: 1.4,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  body: {
    fontSize: '14px',
    fontWeight: 400,
    lineHeight: 1.5,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  bodyMedium: {
    fontSize: '14px',
    fontWeight: 500,
    lineHeight: 1.5,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  bodySemibold: {
    fontSize: '14px',
    fontWeight: 600,
    lineHeight: 1.5,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  small: {
    fontSize: '13px',
    fontWeight: 400,
    lineHeight: 1.4,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  caption: {
    fontSize: '12px',
    fontWeight: 400,
    lineHeight: 1.4,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  captionSemibold: {
    fontSize: '12px',
    fontWeight: 600,
    lineHeight: 1.4,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 14. SPACING SCALE
// Source: CSS Custom Properties (lines 70-78)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlSpacingScale = {
  xxs: '4px',
  xs: '8px',
  s: '12px',
  m: '16px',
  l: '20px',
  xl: '24px',
  xxl: '32px',
  xxxl: '48px',
};

// ════════════════════════════════════════════════════════════════════════════════════
// 15. BORDER RADIUS
// Source: CSS Custom Properties (lines 80-86)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlBorderRadius = {
  none: '0',
  small: '2px',
  medium: '4px',
  large: '6px',
  xlarge: '8px',
  xxlarge: '12px',
  round: '50%',
  pill: '9999px',
};

// ════════════════════════════════════════════════════════════════════════════════════
// 16. SHADOW DEFINITIONS
// Source: CSS Custom Properties (lines 88-93)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlShadows = {
  depth4: '0 1.6px 3.6px rgba(0, 0, 0, 0.13)',
  depth8: '0 3.2px 7.2px rgba(0, 0, 0, 0.13)',
  depth16: '0 6.4px 14.4px rgba(0, 0, 0, 0.13)',
  depth64: '0 12.8px 28.8px rgba(0, 0, 0, 0.13)',
  dropdown: '0 8px 32px rgba(0, 0, 0, 0.2)',
  panel: '-8px 0 32px rgba(0,0,0,0.2)',
};

// ════════════════════════════════════════════════════════════════════════════════════
// 17. BUTTON STYLES
// Source: .btn-* (lines 895-961)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlButtonStyles = {
  /**
   * Base button styles - all buttons inherit these
   * Source: .btn { display: inline-flex; align-items: center; padding: 8px 16px; border-radius: 4px; }
   */
  base: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    padding: '8px 16px',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    border: '1px solid transparent',
    minHeight: '32px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Primary button - blue background, white text
   * Source: .btn-primary { background: var(--jml-primary); color: white; }
   */
  primary: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    padding: '8px 16px',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    minHeight: '32px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    backgroundColor: '#0078d4',
    color: '#ffffff',
    border: '1px solid #0078d4',
    ':hover': {
      backgroundColor: '#106ebe',
      borderColor: '#106ebe',
    },
    ':disabled': {
      opacity: 0.5,
      cursor: 'not-allowed',
    },
  },

  /**
   * Secondary button - white background, gray border
   * Source: .btn-secondary { background: white; color: neutralPrimary; border-color: neutralTertiary; }
   */
  secondary: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    padding: '8px 16px',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    minHeight: '32px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    backgroundColor: '#ffffff',
    color: '#323130',
    border: '1px solid #8a8886',
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
    ':disabled': {
      opacity: 0.5,
      cursor: 'not-allowed',
    },
  },

  /**
   * Tertiary button - transparent, blue text
   * Source: .btn-tertiary { background: transparent; color: primary; border-color: transparent; }
   */
  tertiary: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    padding: '8px 16px',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    minHeight: '32px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    backgroundColor: 'transparent',
    color: '#0078d4',
    border: '1px solid transparent',
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
    ':disabled': {
      opacity: 0.5,
      cursor: 'not-allowed',
    },
  },

  /**
   * Teal button - teal background, white text (CTAs)
   * Source: .btn-teal { background: actionTeal; color: white; }
   */
  teal: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    padding: '8px 16px',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    minHeight: '32px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    backgroundColor: '#03787C',
    color: '#ffffff',
    border: '1px solid #03787C',
    ':hover': {
      backgroundColor: '#026569',
      borderColor: '#026569',
    },
    ':disabled': {
      opacity: 0.5,
      cursor: 'not-allowed',
    },
  },

  /**
   * Danger button - red background, white text
   * Source: .btn-danger { background: error; color: white; }
   */
  danger: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    padding: '8px 16px',
    borderRadius: '4px',
    fontSize: '14px',
    fontWeight: 600,
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    minHeight: '32px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    backgroundColor: '#d13438',
    color: '#ffffff',
    border: '1px solid #d13438',
    ':hover': {
      backgroundColor: '#a52a2d',
      borderColor: '#a52a2d',
    },
    ':disabled': {
      opacity: 0.5,
      cursor: 'not-allowed',
    },
  },

  /**
   * Icon button - square, no background
   * Source: .jml-icon-button { width: 32px; height: 32px; border-radius: 4px; }
   */
  icon: {
    width: '32px',
    height: '32px',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 0,
    borderRadius: '4px',
    fontSize: '16px',
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    backgroundColor: 'transparent',
    color: '#605e5c',
    border: 'none',
    ':hover': {
      backgroundColor: '#f3f2f1',
      color: '#323130',
    },
    ':disabled': {
      opacity: 0.5,
      cursor: 'not-allowed',
    },
  },

  /**
   * Icon button small
   */
  iconSmall: {
    width: '28px',
    height: '28px',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 0,
    borderRadius: '4px',
    fontSize: '14px',
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    backgroundColor: 'transparent',
    color: '#605e5c',
    border: 'none',
    ':hover': {
      backgroundColor: '#f3f2f1',
      color: '#323130',
    },
  },

  /**
   * Icon button with circle background
   */
  iconCircle: {
    width: '36px',
    height: '36px',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 0,
    borderRadius: '50%',
    fontSize: '16px',
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    backgroundColor: '#f3f2f1',
    color: '#605e5c',
    border: 'none',
    ':hover': {
      backgroundColor: '#edebe9',
      color: '#323130',
    },
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 18. BADGE STYLES
// Source: .badge-* (lines 966-995)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlBadgeStyles = {
  /**
   * Base badge styles
   * Source: .badge { display: inline-flex; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
   */
  base: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  // Status badges
  pending: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#fff4ce',
    color: '#986f0b',
  },

  active: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#deecf9',
    color: '#0078d4',
  },

  inProgress: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#deecf9',
    color: '#0078d4',
  },

  completed: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#dff6dd',
    color: '#107c10',
  },

  overdue: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#fde7e9',
    color: '#d13438',
  },

  blocked: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#f3f2f1',
    color: '#605e5c',
  },

  cancelled: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#f3f2f1',
    color: '#8a8886',
    textDecoration: 'line-through',
  },

  // Process type badges
  joiner: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#deecf9',
    color: '#0078d4',
  },

  mover: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#f3e5f5',
    color: '#7b1fa2',
  },

  leaver: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#fff3e0',
    color: '#e65100',
  },

  // Priority badges
  priorityHigh: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#fde7e9',
    color: '#d13438',
  },

  priorityMedium: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#fff4ce',
    color: '#986f0b',
  },

  priorityLow: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '4px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
    backgroundColor: '#dff6dd',
    color: '#107c10',
  },

  // Pill variant (larger padding)
  pill: {
    padding: '6px 14px',
    borderRadius: '16px',
    fontSize: '13px',
  },

  // Count badge (for notification counts)
  count: {
    minWidth: '20px',
    height: '20px',
    padding: '0 6px',
    borderRadius: '10px',
    fontSize: '11px',
    fontWeight: 600,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 19. PANEL STYLES (SharePoint Fly-in Panel)
// Source: .panel-* (lines 1057-1128)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlPanelStyles = {
  /**
   * Panel backdrop overlay
   * Source: .panel-backdrop { position: absolute; inset: 0; background: rgba(0,0,0,0.4); }
   */
  backdrop: {
    position: 'fixed' as const,
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0, 0, 0, 0.4)',
    zIndex: 1000,
  },

  /**
   * Panel container base
   * Source: .panel-mock { position: absolute; right: 0; background: white; box-shadow: -8px 0 32px rgba(0,0,0,0.2); }
   */
  container: {
    position: 'fixed' as const,
    right: 0,
    top: 0,
    bottom: 0,
    backgroundColor: '#ffffff',
    boxShadow: '-8px 0 32px rgba(0, 0, 0, 0.2)',
    display: 'flex',
    flexDirection: 'column' as const,
    zIndex: 1001,
  },

  // Panel sizes
  small: { width: '340px' },
  medium: { width: '480px' },
  large: { width: '640px' },
  extraLarge: { width: '800px' },

  /**
   * Panel header
   * Source: .panel-header { padding: 16px 24px; border-bottom: 1px solid; display: flex; justify-content: space-between; }
   */
  header: {
    padding: '16px 24px',
    borderBottom: '1px solid #edebe9',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    backgroundColor: '#ffffff',
    flexShrink: 0,
  },

  /**
   * Panel header title
   * Source: .panel-header-title { font-size: 20px; font-weight: 600; }
   */
  headerTitle: {
    fontSize: '20px',
    fontWeight: 600,
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Panel close button
   * Source: .panel-close-btn { width: 32px; height: 32px; border: none; background: transparent; }
   */
  closeButton: {
    width: '32px',
    height: '32px',
    border: 'none',
    backgroundColor: 'transparent',
    cursor: 'pointer',
    borderRadius: '4px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#605e5c',
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
  },

  /**
   * Panel content area
   * Source: .panel-content { flex: 1; padding: 24px; overflow-y: auto; }
   */
  content: {
    flex: 1,
    padding: '24px',
    overflowY: 'auto' as const,
    backgroundColor: '#ffffff',
  },

  /**
   * Panel footer with actions
   * Source: .panel-footer { padding: 16px 24px; border-top: 1px solid; display: flex; justify-content: flex-end; gap: 8px; }
   */
  footer: {
    padding: '16px 24px',
    borderTop: '1px solid #edebe9',
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    backgroundColor: '#ffffff',
    flexShrink: 0,
  },

  /**
   * Panel section within content
   */
  section: {
    marginBottom: '24px',
  },

  sectionTitle: {
    fontSize: '16px',
    fontWeight: 600,
    color: '#323130',
    marginBottom: '12px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 20. MODAL / DIALOG STYLES
// Source: .modal-* (lines 1131-1197)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlModalStyles = {
  /**
   * Modal backdrop overlay
   * Source: .modal-backdrop { position: absolute; inset: 0; background: rgba(0,0,0,0.4); }
   */
  backdrop: {
    position: 'fixed' as const,
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0, 0, 0, 0.4)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 1000,
  },

  /**
   * Modal container
   * Source: .modal-mock { position: relative; background: white; border-radius: 8px; box-shadow: depth64; }
   */
  container: {
    position: 'relative' as const,
    backgroundColor: '#ffffff',
    borderRadius: '8px',
    boxShadow: '0 12.8px 28.8px rgba(0, 0, 0, 0.13)',
    maxWidth: '90%',
    maxHeight: '90vh',
    display: 'flex',
    flexDirection: 'column' as const,
    zIndex: 1001,
  },

  // Modal sizes
  small: { width: '340px' },
  medium: { width: '480px' },
  large: { width: '640px' },

  /**
   * Modal header
   * Source: .modal-header { padding: 20px 24px 12px; background: white; }
   */
  header: {
    padding: '20px 24px 12px',
    backgroundColor: '#ffffff',
    borderRadius: '8px 8px 0 0',
  },

  /**
   * Modal title
   * Source: .modal-title { font-size: 20px; font-weight: 600; color: neutralPrimary; }
   */
  title: {
    fontSize: '20px',
    fontWeight: 600,
    color: '#323130',
    marginBottom: '8px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Modal subtext
   * Source: .modal-subtext { font-size: 14px; color: neutralSecondary; }
   */
  subtext: {
    fontSize: '14px',
    color: '#605e5c',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Modal content area
   * Source: .modal-content { padding: 0 24px 24px; background: white; }
   */
  content: {
    padding: '0 24px 24px',
    backgroundColor: '#ffffff',
    overflowY: 'auto' as const,
    flex: 1,
  },

  /**
   * Modal footer with actions
   * Source: .modal-footer { padding: 16px 24px; border-top: 1px solid; display: flex; justify-content: flex-end; gap: 8px; }
   */
  footer: {
    padding: '16px 24px',
    borderTop: '1px solid #edebe9',
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    backgroundColor: '#ffffff',
    borderRadius: '0 0 8px 8px',
  },

  /**
   * Modal icon (for confirmation dialogs)
   * Source: .modal-icon { font-size: 24px; margin-bottom: 12px; }
   */
  icon: {
    fontSize: '24px',
    marginBottom: '12px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },

  iconWarning: { color: '#ffb900' },
  iconError: { color: '#d13438' },
  iconInfo: { color: '#0078d4' },
  iconSuccess: { color: '#107c10' },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 21. DROPDOWN STYLES
// Source: .dropdown-* (lines 1320-1351)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlDropdownStyles = {
  /**
   * Dropdown trigger button
   */
  trigger: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '8px 12px',
    backgroundColor: '#ffffff',
    border: '1px solid #c8c6c4',
    borderRadius: '4px',
    fontSize: '14px',
    color: '#323130',
    cursor: 'pointer',
    minWidth: '150px',
    transition: 'all 0.1s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      borderColor: '#8a8886',
    },
    ':focus': {
      borderColor: '#0078d4',
      outline: 'none',
    },
  },

  /**
   * Dropdown callout/panel
   * Source: .dropdown-callout { position: absolute; background: white; border: 1px solid; border-radius: 4px; box-shadow: dropdown; }
   */
  callout: {
    position: 'absolute' as const,
    backgroundColor: '#ffffff',
    border: '1px solid #edebe9',
    borderRadius: '4px',
    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.2)',
    minWidth: '150px',
    maxHeight: '300px',
    overflowY: 'auto' as const,
    zIndex: 1000,
  },

  /**
   * Dropdown item
   * Source: .dropdown-item { padding: 10px 12px; cursor: pointer; }
   */
  item: {
    padding: '10px 12px',
    cursor: 'pointer',
    fontSize: '14px',
    color: '#323130',
    transition: 'background-color 0.1s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
  },

  itemSelected: {
    padding: '10px 12px',
    cursor: 'pointer',
    fontSize: '14px',
    color: '#0078d4',
    backgroundColor: '#deecf9',
    fontWeight: 600,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  itemDisabled: {
    padding: '10px 12px',
    cursor: 'not-allowed',
    fontSize: '14px',
    color: '#a19f9d',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Dropdown divider
   */
  divider: {
    height: '1px',
    backgroundColor: '#edebe9',
    margin: '4px 0',
  },

  /**
   * Dropdown header (for grouped items)
   */
  header: {
    padding: '8px 12px',
    fontSize: '12px',
    fontWeight: 600,
    color: '#605e5c',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 22. MESSAGE BAR STYLES
// Source: .message-bar-* (lines 1200-1226)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlMessageBarStyles = {
  /**
   * Base message bar container
   * Source: .message-bar { display: flex; align-items: center; gap: 12px; padding: 12px 16px; border-radius: 0 4px 4px 0; }
   */
  base: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px 16px',
    borderRadius: '0 4px 4px 0',
    marginBottom: '12px',
  },

  icon: {
    fontSize: '16px',
    flexShrink: 0,
  },

  content: {
    flex: 1,
    fontSize: '14px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  dismiss: {
    backgroundColor: 'transparent',
    border: 'none',
    cursor: 'pointer',
    padding: '4px',
    opacity: 0.7,
    ':hover': {
      opacity: 1,
    },
  },

  /**
   * Info message bar (blue)
   * Source: .message-bar.info { background: #deecf9; border-left: 4px solid primary; }
   */
  info: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px 16px',
    borderRadius: '0 4px 4px 0',
    marginBottom: '12px',
    backgroundColor: '#deecf9',
    borderLeft: '4px solid #0078d4',
  },

  infoIcon: {
    color: '#0078d4',
    fontSize: '16px',
    flexShrink: 0,
  },

  /**
   * Success message bar (green)
   * Source: .message-bar.success { background: #dff6dd; border-left: 4px solid success; }
   */
  success: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px 16px',
    borderRadius: '0 4px 4px 0',
    marginBottom: '12px',
    backgroundColor: '#dff6dd',
    borderLeft: '4px solid #107c10',
  },

  successIcon: {
    color: '#107c10',
    fontSize: '16px',
    flexShrink: 0,
  },

  /**
   * Warning message bar (yellow)
   * Source: .message-bar.warning { background: #fff4ce; border-left: 4px solid warning; }
   */
  warning: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px 16px',
    borderRadius: '0 4px 4px 0',
    marginBottom: '12px',
    backgroundColor: '#fff4ce',
    borderLeft: '4px solid #ffb900',
  },

  warningIcon: {
    color: '#986f0b',
    fontSize: '16px',
    flexShrink: 0,
  },

  /**
   * Error message bar (red)
   * Source: .message-bar.error { background: #fde7e9; border-left: 4px solid error; }
   */
  error: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px 16px',
    borderRadius: '0 4px 4px 0',
    marginBottom: '12px',
    backgroundColor: '#fde7e9',
    borderLeft: '4px solid #d13438',
  },

  errorIcon: {
    color: '#d13438',
    fontSize: '16px',
    flexShrink: 0,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 23. FORM STYLES
// Source: .form-* (lines 1229-1265)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlFormStyles = {
  /**
   * Form group container
   */
  group: {
    marginBottom: '16px',
  },

  /**
   * Form label
   */
  label: {
    display: 'block',
    marginBottom: '4px',
    fontSize: '14px',
    fontWeight: 600,
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  labelRequired: {
    '::after': {
      content: '" *"',
      color: '#d13438',
    },
  },

  /**
   * Text input field
   */
  input: {
    width: '100%',
    padding: '8px 12px',
    fontSize: '14px',
    border: '1px solid #c8c6c4',
    borderRadius: '4px',
    backgroundColor: '#ffffff',
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    transition: 'border-color 0.1s ease',
    ':focus': {
      borderColor: '#0078d4',
      outline: 'none',
    },
    ':disabled': {
      backgroundColor: '#f3f2f1',
      color: '#a19f9d',
      cursor: 'not-allowed',
    },
  },

  inputError: {
    borderColor: '#d13438',
  },

  /**
   * Textarea
   */
  textarea: {
    width: '100%',
    padding: '8px 12px',
    fontSize: '14px',
    border: '1px solid #c8c6c4',
    borderRadius: '4px',
    backgroundColor: '#ffffff',
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    resize: 'vertical' as const,
    minHeight: '80px',
    transition: 'border-color 0.1s ease',
    ':focus': {
      borderColor: '#0078d4',
      outline: 'none',
    },
  },

  /**
   * Helper text below input
   */
  helperText: {
    marginTop: '4px',
    fontSize: '12px',
    color: '#605e5c',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Error message below input
   */
  errorText: {
    marginTop: '4px',
    fontSize: '12px',
    color: '#d13438',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Form row (horizontal layout)
   */
  row: {
    display: 'flex',
    gap: '16px',
    flexWrap: 'wrap' as const,
  },

  /**
   * Form column
   */
  col: {
    flex: 1,
    minWidth: '200px',
  },

  colHalf: {
    flex: '0 0 calc(50% - 8px)',
    minWidth: '200px',
  },

  colThird: {
    flex: '0 0 calc(33.333% - 11px)',
    minWidth: '150px',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 24. NAVIGATION BAR STYLES
// Source: .jml-header, .jml-nav-bar (lines 193-305)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlNavigationStyles = {
  /**
   * Top navigation bar container
   * Source: .jml-header { background: linear-gradient(135deg, #0078d4 0%, #004578 100%); }
   */
  header: {
    background: 'linear-gradient(135deg, #0078d4 0%, #004578 100%)',
    color: '#ffffff',
    position: 'relative' as const,
    zIndex: 100000,
  },

  /**
   * Nav bar inner container
   * Source: .jml-nav-bar { height: 56px; padding: 8px 24px; }
   */
  navBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '8px 24px',
    height: '56px',
    boxShadow: '0 2px 8px rgba(0, 0, 0, 0.2)',
  },

  /**
   * Brand logo and text
   */
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    textDecoration: 'none',
    color: '#ffffff',
  },

  brandLogo: {
    width: '32px',
    height: '32px',
    backgroundColor: '#ffffff',
    borderRadius: '6px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 700,
    fontSize: '14px',
    color: '#0078d4',
  },

  brandText: {
    fontSize: '16px',
    fontWeight: 600,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  brandTextLight: {
    opacity: 0.8,
    fontWeight: 400,
  },

  /**
   * Nav container
   */
  nav: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },

  /**
   * Nav link/button
   * Source: .jml-nav-link { height: 40px; font-size: 13px; font-weight: 500; }
   */
  navLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    padding: '0 16px',
    height: '40px',
    color: 'rgba(255, 255, 255, 0.85)',
    textDecoration: 'none',
    fontSize: '13px',
    fontWeight: 500,
    borderRadius: '4px',
    border: 'none',
    backgroundColor: 'transparent',
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      color: '#ffffff',
      backgroundColor: 'rgba(255, 255, 255, 0.1)',
    },
  },

  navLinkActive: {
    color: '#ffffff',
    backgroundColor: 'rgba(255, 255, 255, 0.15)',
  },

  /**
   * User section (right side)
   */
  userSection: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },

  /**
   * Icon button (search, notifications)
   * Source: .jml-icon-button { width: 36px; height: 36px; border-radius: 50%; }
   */
  iconButton: {
    width: '36px',
    height: '36px',
    borderRadius: '50%',
    backgroundColor: 'rgba(255, 255, 255, 0.1)',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'background 0.15s ease',
    ':hover': {
      backgroundColor: 'rgba(255, 255, 255, 0.2)',
    },
  },

  /**
   * Page header block (below nav)
   * Source: .jml-page-header { background: linear-gradient(180deg, #004578 0%, #005a9e 100%); }
   */
  pageHeader: {
    padding: '24px 40px 24px 72px',
    background: 'linear-gradient(180deg, #004578 0%, #005a9e 100%)',
    borderTop: '1px solid rgba(255, 255, 255, 0.2)',
    borderBottom: '4px solid #0078d4',
    minHeight: '100px',
    color: '#ffffff',
  },

  pageTitle: {
    fontSize: '28px',
    fontWeight: 600,
    margin: '0 0 8px 0',
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    textShadow: '0 1px 2px rgba(0, 0, 0, 0.2)',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  pageTitleIcon: {
    fontSize: '32px',
  },

  pageDescription: {
    fontSize: '14px',
    color: 'rgba(255, 255, 255, 0.9)',
    margin: 0,
    maxWidth: '600px',
    lineHeight: 1.5,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 25. BREADCRUMB STYLES
// Source: .jml-breadcrumb (lines 321-345)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlBreadcrumbStyles = {
  /**
   * Breadcrumb container
   */
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '13px',
    marginBottom: '8px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Light theme (for dark backgrounds like page header)
   */
  containerLight: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '13px',
    marginBottom: '8px',
    color: 'rgba(255, 255, 255, 0.8)',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Breadcrumb link
   */
  link: {
    color: '#0078d4',
    textDecoration: 'none',
    ':hover': {
      textDecoration: 'underline',
    },
  },

  linkLight: {
    color: 'rgba(255, 255, 255, 0.8)',
    textDecoration: 'none',
    ':hover': {
      color: '#ffffff',
      textDecoration: 'underline',
    },
  },

  /**
   * Separator between items
   */
  separator: {
    color: '#a19f9d',
  },

  separatorLight: {
    color: 'rgba(255, 255, 255, 0.5)',
  },

  /**
   * Current page (last item)
   */
  current: {
    color: '#323130',
    fontWeight: 500,
  },

  currentLight: {
    color: '#ffffff',
    fontWeight: 500,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 26. TAB PANEL STYLES
// Source: .jml-tab-panel, .jml-tab-button (lines 375-425)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlTabPanelStyles = {
  /**
   * Tab panel container - VISUAL STYLING ONLY (no margins)
   *
   * USE THIS FOR: Components that need tab panel APPEARANCE without positioning.
   *
   * FOR ACTUAL TAB PANELS: Use TabPanelStyles.ts → useTabPanelStyles()
   * That hook includes the 24px margins for proper page layout.
   *
   * Example:
   *   import { useTabPanelStyles } from '../styles/TabPanelStyles';
   *   const tabStyles = useTabPanelStyles();
   *   <div className={tabStyles.tabPanel}>...</div>
   */
  container: {
    backgroundColor: '#ffffff',
    border: '1px solid #edebe9',
    borderLeft: '4px solid #0078d4',
    borderRadius: '8px',
    padding: '8px 12px 8px 16px',
    boxShadow: '0 1px 3px rgba(0, 0, 0, 0.08)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
    flexWrap: 'wrap' as const,
    // NO MARGINS - use TabPanelStyles.ts (useTabPanelStyles) for positioned tab panels
  },

  /**
   * Tab button (inactive)
   * Source: .jml-tab-button { padding: 10px 16px; border-radius: 6px; font-size: 14px; font-weight: 500; }
   */
  tab: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 16px',
    backgroundColor: 'transparent',
    border: 'none',
    borderRadius: '6px',
    fontSize: '14px',
    fontWeight: 500,
    color: '#323130',
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
  },

  /**
   * Tab button (active)
   * Source: .jml-tab-button.active { background: #0078d4; color: white; font-weight: 600; }
   */
  tabActive: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 16px',
    backgroundColor: '#0078d4',
    border: 'none',
    borderRadius: '6px',
    fontSize: '14px',
    fontWeight: 600,
    color: '#ffffff',
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Tab icon styling
   */
  icon: {
    width: '18px',
    height: '18px',
    strokeWidth: 1.5,
  },

  iconInactive: {
    color: '#0078d4',
  },

  iconActive: {
    color: '#ffffff',
  },

  /**
   * Tab content panel (below tabs)
   */
  content: {
    backgroundColor: '#ffffff',
    padding: '24px',
    borderRadius: '0 0 8px 8px',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 27. FOOTER STYLES
// Source: .jml-footer (lines 556-645)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlFooterStyles = {
  /**
   * Footer container
   * Source: .jml-footer { background: linear-gradient(135deg, #004578 0%, #0078d4 100%); }
   */
  container: {
    background: 'linear-gradient(135deg, #004578 0%, #0078d4 100%)',
    color: '#ffffff',
    marginTop: 'auto',
    flexShrink: 0,
    width: '100%',
  },

  /**
   * Footer inner content
   */
  content: {
    maxWidth: '1400px',
    margin: '0 auto',
    padding: '16px 24px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexWrap: 'wrap' as const,
    gap: '16px',
    minHeight: '60px',
  },

  /**
   * Left section (brand + copyright)
   */
  left: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
  },

  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },

  logo: {
    width: '24px',
    height: '24px',
    backgroundColor: '#ffffff',
    borderRadius: '4px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 700,
    fontSize: '10px',
    color: '#0078d4',
  },

  copyright: {
    fontSize: '12px',
    opacity: 0.85,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Center section (links)
   */
  center: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },

  link: {
    fontSize: '12px',
    color: 'rgba(255, 255, 255, 0.9)',
    textDecoration: 'none',
    transition: 'color 0.15s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      color: '#ffffff',
      textDecoration: 'underline',
    },
  },

  separator: {
    fontSize: '12px',
    color: 'rgba(255, 255, 255, 0.5)',
  },

  /**
   * Right section (version)
   */
  right: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
  },

  version: {
    fontSize: '11px',
    opacity: 0.7,
    padding: '4px 8px',
    backgroundColor: 'rgba(255, 255, 255, 0.1)',
    borderRadius: '4px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 28. LINK STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlLinkStyles = {
  /**
   * Standard link (primary blue)
   */
  default: {
    color: '#0078d4',
    textDecoration: 'none',
    cursor: 'pointer',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      textDecoration: 'underline',
      color: '#106ebe',
    },
  },

  /**
   * Subtle link (neutral color)
   */
  subtle: {
    color: '#605e5c',
    textDecoration: 'none',
    cursor: 'pointer',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      textDecoration: 'underline',
      color: '#323130',
    },
  },

  /**
   * Disabled link
   */
  disabled: {
    color: '#a19f9d',
    textDecoration: 'none',
    cursor: 'not-allowed',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Inverted link (for dark backgrounds)
   */
  inverted: {
    color: 'rgba(255, 255, 255, 0.9)',
    textDecoration: 'none',
    cursor: 'pointer',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      color: '#ffffff',
      textDecoration: 'underline',
    },
  },

  /**
   * Button-styled link
   */
  button: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '6px',
    color: '#0078d4',
    textDecoration: 'none',
    cursor: 'pointer',
    fontWeight: 600,
    fontSize: '14px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      textDecoration: 'underline',
    },
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 29. TOOLTIP STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlTooltipStyles = {
  /**
   * Dark tooltip (default)
   */
  container: {
    backgroundColor: '#323130',
    color: '#ffffff',
    padding: '8px 12px',
    borderRadius: '4px',
    fontSize: '12px',
    fontWeight: 400,
    maxWidth: '280px',
    boxShadow: '0 3.2px 7.2px rgba(0, 0, 0, 0.13)',
    zIndex: 10000,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Light tooltip (inverted)
   */
  containerLight: {
    backgroundColor: '#ffffff',
    color: '#323130',
    padding: '8px 12px',
    borderRadius: '4px',
    fontSize: '12px',
    fontWeight: 400,
    maxWidth: '280px',
    border: '1px solid #edebe9',
    boxShadow: '0 3.2px 7.2px rgba(0, 0, 0, 0.13)',
    zIndex: 10000,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Tooltip title
   */
  title: {
    fontWeight: 600,
    marginBottom: '4px',
    fontSize: '13px',
  },

  /**
   * Tooltip content
   */
  content: {
    fontSize: '12px',
    lineHeight: 1.4,
  },

  /**
   * Tooltip arrow (pointing down)
   */
  arrow: {
    position: 'absolute' as const,
    width: 0,
    height: 0,
    borderLeft: '6px solid transparent',
    borderRight: '6px solid transparent',
    borderTop: '6px solid #323130',
  },

  arrowLight: {
    position: 'absolute' as const,
    width: 0,
    height: 0,
    borderLeft: '6px solid transparent',
    borderRight: '6px solid transparent',
    borderTop: '6px solid #ffffff',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 30. PROGRESS STYLES (BARS + SPINNERS)
// Source: .spinner (lines 1293-1304)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlProgressStyles = {
  /**
   * Progress bar container
   */
  barContainer: {
    width: '100%',
    height: '4px',
    backgroundColor: '#edebe9',
    borderRadius: '2px',
    overflow: 'hidden',
  },

  barContainerLarge: {
    width: '100%',
    height: '8px',
    backgroundColor: '#edebe9',
    borderRadius: '4px',
    overflow: 'hidden',
  },

  /**
   * Progress bar fill
   */
  bar: {
    height: '100%',
    backgroundColor: '#0078d4',
    borderRadius: '2px',
    transition: 'width 0.3s ease',
  },

  barSuccess: {
    height: '100%',
    backgroundColor: '#107c10',
    borderRadius: '2px',
    transition: 'width 0.3s ease',
  },

  barWarning: {
    height: '100%',
    backgroundColor: '#ffb900',
    borderRadius: '2px',
    transition: 'width 0.3s ease',
  },

  barError: {
    height: '100%',
    backgroundColor: '#d13438',
    borderRadius: '2px',
    transition: 'width 0.3s ease',
  },

  /**
   * Progress label
   */
  label: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '8px',
    fontSize: '14px',
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  percentage: {
    fontSize: '14px',
    fontWeight: 600,
    color: '#0078d4',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Spinner (loading indicator)
   * Source: .spinner { width: 32px; height: 32px; border: 3px solid neutralLight; border-top-color: primary; }
   */
  spinner: {
    width: '32px',
    height: '32px',
    border: '3px solid #edebe9',
    borderTopColor: '#0078d4',
    borderRadius: '50%',
    animation: 'spin 1s linear infinite',
  },

  spinnerSmall: {
    width: '20px',
    height: '20px',
    border: '2px solid #edebe9',
    borderTopColor: '#0078d4',
    borderRadius: '50%',
    animation: 'spin 1s linear infinite',
  },

  spinnerLarge: {
    width: '48px',
    height: '48px',
    border: '4px solid #edebe9',
    borderTopColor: '#0078d4',
    borderRadius: '50%',
    animation: 'spin 1s linear infinite',
  },

  /**
   * Loading overlay
   */
  overlay: {
    position: 'absolute' as const,
    inset: 0,
    backgroundColor: 'rgba(255, 255, 255, 0.8)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexDirection: 'column' as const,
    gap: '16px',
    zIndex: 100,
  },

  loadingText: {
    fontSize: '14px',
    color: '#605e5c',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 31. CHECKBOX STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlCheckboxStyles = {
  /**
   * Checkbox container (label + input)
   */
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    cursor: 'pointer',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Checkbox input (visual box)
   */
  input: {
    width: '20px',
    height: '20px',
    border: '1px solid #605e5c',
    borderRadius: '2px',
    backgroundColor: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'all 0.1s ease',
    flexShrink: 0,
  },

  inputChecked: {
    width: '20px',
    height: '20px',
    border: '1px solid #0078d4',
    borderRadius: '2px',
    backgroundColor: '#0078d4',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'all 0.1s ease',
    flexShrink: 0,
  },

  inputDisabled: {
    width: '20px',
    height: '20px',
    border: '1px solid #c8c6c4',
    borderRadius: '2px',
    backgroundColor: '#f3f2f1',
    cursor: 'not-allowed',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },

  /**
   * Checkmark icon
   */
  checkmark: {
    color: '#ffffff',
    fontSize: '14px',
  },

  /**
   * Checkbox label text
   */
  label: {
    fontSize: '14px',
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  labelDisabled: {
    fontSize: '14px',
    color: '#a19f9d',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 32. RADIO BUTTON STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlRadioStyles = {
  /**
   * Radio container
   */
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    cursor: 'pointer',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Radio input (visual circle)
   */
  input: {
    width: '20px',
    height: '20px',
    border: '1px solid #605e5c',
    borderRadius: '50%',
    backgroundColor: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'all 0.1s ease',
    flexShrink: 0,
  },

  inputChecked: {
    width: '20px',
    height: '20px',
    border: '1px solid #0078d4',
    borderRadius: '50%',
    backgroundColor: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'all 0.1s ease',
    flexShrink: 0,
  },

  inputDisabled: {
    width: '20px',
    height: '20px',
    border: '1px solid #c8c6c4',
    borderRadius: '50%',
    backgroundColor: '#f3f2f1',
    cursor: 'not-allowed',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },

  /**
   * Radio dot (inner circle when selected)
   */
  dot: {
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    backgroundColor: '#0078d4',
  },

  dotDisabled: {
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    backgroundColor: '#a19f9d',
  },

  /**
   * Radio label text
   */
  label: {
    fontSize: '14px',
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  labelDisabled: {
    fontSize: '14px',
    color: '#a19f9d',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Radio group
   */
  group: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '12px',
  },

  groupHorizontal: {
    display: 'flex',
    flexDirection: 'row' as const,
    gap: '24px',
    flexWrap: 'wrap' as const,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 33. TOGGLE / SWITCH STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlToggleStyles = {
  /**
   * Toggle container
   */
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    cursor: 'pointer',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Toggle track (background)
   */
  track: {
    width: '44px',
    height: '22px',
    backgroundColor: '#605e5c',
    borderRadius: '11px',
    position: 'relative' as const,
    transition: 'background-color 0.2s ease',
    cursor: 'pointer',
  },

  trackOn: {
    width: '44px',
    height: '22px',
    backgroundColor: '#0078d4',
    borderRadius: '11px',
    position: 'relative' as const,
    transition: 'background-color 0.2s ease',
    cursor: 'pointer',
  },

  trackDisabled: {
    width: '44px',
    height: '22px',
    backgroundColor: '#c8c6c4',
    borderRadius: '11px',
    position: 'relative' as const,
    cursor: 'not-allowed',
  },

  /**
   * Toggle thumb (circle)
   */
  thumb: {
    width: '18px',
    height: '18px',
    backgroundColor: '#ffffff',
    borderRadius: '50%',
    position: 'absolute' as const,
    top: '2px',
    left: '2px',
    transition: 'left 0.2s ease',
    boxShadow: '0 1px 3px rgba(0, 0, 0, 0.2)',
  },

  thumbOn: {
    width: '18px',
    height: '18px',
    backgroundColor: '#ffffff',
    borderRadius: '50%',
    position: 'absolute' as const,
    top: '2px',
    left: '24px',
    transition: 'left 0.2s ease',
    boxShadow: '0 1px 3px rgba(0, 0, 0, 0.2)',
  },

  /**
   * Toggle label
   */
  label: {
    fontSize: '14px',
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  labelDisabled: {
    fontSize: '14px',
    color: '#a19f9d',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * On/Off text labels
   */
  stateLabel: {
    fontSize: '12px',
    color: '#605e5c',
    fontWeight: 500,
    minWidth: '24px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 34. AVATAR / PERSONA STYLES
// Source: .jml-avatar (lines 294-305)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlAvatarStyles = {
  /**
   * Avatar sizes
   */
  small: {
    width: '24px',
    height: '24px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '10px',
    fontWeight: 600,
    color: '#ffffff',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  medium: {
    width: '32px',
    height: '32px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '13px',
    fontWeight: 600,
    color: '#ffffff',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  large: {
    width: '48px',
    height: '48px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '18px',
    fontWeight: 600,
    color: '#ffffff',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  xlarge: {
    width: '72px',
    height: '72px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '28px',
    fontWeight: 600,
    color: '#ffffff',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Avatar colors (for initials-based avatars)
   */
  colorBlue: { backgroundColor: '#0078d4' },
  colorTeal: { backgroundColor: '#03787C' },
  colorPurple: { backgroundColor: '#8764b8' },
  colorGreen: { backgroundColor: '#107c10' },
  colorOrange: { backgroundColor: '#f57c00' },
  colorRed: { backgroundColor: '#d13438' },
  colorGray: { backgroundColor: '#605e5c' },

  /**
   * Persona card (avatar + name + details)
   */
  personaCard: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },

  personaDetails: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '2px',
  },

  personaName: {
    fontSize: '14px',
    fontWeight: 600,
    color: '#323130',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  personaSecondary: {
    fontSize: '12px',
    color: '#605e5c',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Avatar with image
   */
  image: {
    objectFit: 'cover' as const,
  },

  /**
   * Avatar with presence indicator
   */
  presence: {
    position: 'absolute' as const,
    bottom: 0,
    right: 0,
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    border: '2px solid #ffffff',
  },

  presenceOnline: { backgroundColor: '#107c10' },
  presenceAway: { backgroundColor: '#ffb900' },
  presenceBusy: { backgroundColor: '#d13438' },
  presenceOffline: { backgroundColor: '#8a8886' },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 35. RESIZABLE TABLE STYLES (with column resizing)
// User requirement: tables must all have column resizing
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlTableResizableStyles = {
  /**
   * Table wrapper with horizontal scroll support
   */
  wrapper: {
    width: '100%',
    overflowX: 'auto' as const,
    border: '1px solid #edebe9',
    borderRadius: '8px',
    backgroundColor: '#ffffff',
  },

  /**
   * Table with fixed layout for resizing
   */
  table: {
    width: '100%',
    minWidth: '600px',
    borderCollapse: 'collapse' as const,
    tableLayout: 'fixed' as const,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Table header with blue accent
   */
  thead: {
    backgroundColor: '#f3f2f1',
    position: 'relative' as const,
  },

  theadWithAccent: {
    backgroundColor: '#f3f2f1',
    position: 'relative' as const,
    '::before': {
      content: '""',
      position: 'absolute' as const,
      left: 0,
      top: 0,
      bottom: 0,
      width: '4px',
      backgroundColor: '#0078d4',
    },
  },

  /**
   * Header cell with resize handle
   */
  th: {
    position: 'relative' as const,
    padding: '12px 16px',
    textAlign: 'left' as const,
    fontSize: '12px',
    fontWeight: 600,
    color: '#323130',
    borderBottom: '1px solid #edebe9',
    userSelect: 'none' as const,
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },

  thFirst: {
    position: 'relative' as const,
    padding: '12px 16px 12px 20px',
    textAlign: 'left' as const,
    fontSize: '12px',
    fontWeight: 600,
    color: '#323130',
    borderBottom: '1px solid #edebe9',
    userSelect: 'none' as const,
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },

  /**
   * Resize handle (vertical bar between columns)
   */
  resizeHandle: {
    position: 'absolute' as const,
    right: 0,
    top: 0,
    bottom: 0,
    width: '5px',
    cursor: 'col-resize',
    backgroundColor: 'transparent',
    transition: 'background-color 0.1s ease',
    ':hover': {
      backgroundColor: '#0078d4',
    },
  },

  resizeHandleActive: {
    position: 'absolute' as const,
    right: 0,
    top: 0,
    bottom: 0,
    width: '5px',
    cursor: 'col-resize',
    backgroundColor: '#0078d4',
  },

  /**
   * Table body cell
   */
  td: {
    padding: '12px 16px',
    borderBottom: '1px solid #edebe9',
    fontSize: '14px',
    color: '#323130',
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },

  tdFirst: {
    padding: '12px 16px 12px 20px',
    borderBottom: '1px solid #edebe9',
    fontSize: '14px',
    color: '#323130',
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },

  /**
   * Row hover
   */
  tr: {
    transition: 'background-color 0.15s ease',
    ':hover': {
      backgroundColor: '#eff6fc',
    },
  },

  trLast: {
    transition: 'background-color 0.15s ease',
    ':hover': {
      backgroundColor: '#eff6fc',
    },
  },

  /**
   * Sortable header
   */
  thSortable: {
    cursor: 'pointer',
    ':hover': {
      backgroundColor: '#edebe9',
    },
  },

  sortIcon: {
    marginLeft: '8px',
    fontSize: '10px',
    color: '#605e5c',
  },

  sortIconActive: {
    marginLeft: '8px',
    fontSize: '10px',
    color: '#0078d4',
  },

  /**
   * Column minimum widths by type
   */
  columnWidths: {
    checkbox: '40px',
    icon: '48px',
    status: '100px',
    date: '120px',
    name: '180px',
    email: '220px',
    actions: '100px',
    default: '150px',
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 36. SLIDER / RANGE STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlSliderStyles = {
  /**
   * Slider container
   */
  container: {
    width: '100%',
    padding: '8px 0',
  },

  /**
   * Track (background bar)
   */
  track: {
    width: '100%',
    height: '4px',
    backgroundColor: '#edebe9',
    borderRadius: '2px',
    position: 'relative' as const,
  },

  /**
   * Filled portion of track
   */
  fill: {
    height: '100%',
    backgroundColor: '#0078d4',
    borderRadius: '2px',
    position: 'absolute' as const,
    left: 0,
    top: 0,
  },

  /**
   * Thumb (draggable circle)
   */
  thumb: {
    width: '16px',
    height: '16px',
    backgroundColor: '#0078d4',
    borderRadius: '50%',
    position: 'absolute' as const,
    top: '50%',
    transform: 'translateY(-50%)',
    cursor: 'pointer',
    boxShadow: '0 1px 3px rgba(0, 0, 0, 0.2)',
    transition: 'transform 0.1s ease',
    ':hover': {
      transform: 'translateY(-50%) scale(1.1)',
    },
    ':active': {
      transform: 'translateY(-50%) scale(1.2)',
    },
  },

  thumbDisabled: {
    width: '16px',
    height: '16px',
    backgroundColor: '#c8c6c4',
    borderRadius: '50%',
    position: 'absolute' as const,
    top: '50%',
    transform: 'translateY(-50%)',
    cursor: 'not-allowed',
  },

  /**
   * Value label
   */
  valueLabel: {
    fontSize: '14px',
    fontWeight: 600,
    color: '#323130',
    marginBottom: '8px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Min/Max labels
   */
  rangeLabels: {
    display: 'flex',
    justifyContent: 'space-between',
    marginTop: '8px',
    fontSize: '12px',
    color: '#605e5c',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 37. DIVIDER / SEPARATOR STYLES
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlDividerStyles = {
  /**
   * Horizontal divider
   */
  horizontal: {
    width: '100%',
    height: '1px',
    backgroundColor: '#edebe9',
    margin: '16px 0',
  },

  horizontalLight: {
    width: '100%',
    height: '1px',
    backgroundColor: '#f3f2f1',
    margin: '16px 0',
  },

  horizontalDark: {
    width: '100%',
    height: '1px',
    backgroundColor: '#c8c6c4',
    margin: '16px 0',
  },

  /**
   * Vertical divider
   */
  vertical: {
    width: '1px',
    height: '100%',
    backgroundColor: '#edebe9',
    margin: '0 16px',
  },

  /**
   * Divider with text
   */
  withText: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
    margin: '24px 0',
  },

  withTextLine: {
    flex: 1,
    height: '1px',
    backgroundColor: '#edebe9',
  },

  withTextLabel: {
    fontSize: '12px',
    fontWeight: 600,
    color: '#605e5c',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    whiteSpace: 'nowrap' as const,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// UNIFIED EXPORT - ALL JML STYLES
// ════════════════════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════════════════════
// 38. NAV ICON CLASSIFICATION STYLES (NAV-APP)
// Source: Navigation Icon Classification section
// The JML navigation bar uses four distinct icon categories with specific visibility rules
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlNavIconStyles = {
  /**
   * Icon categories with their standard colors
   */
  categories: {
    coreProcesses: { color: '#0078d4', label: 'Core JML Processes' },
    enterprise: { color: '#03787C', label: 'Enterprise Premium' },
    roleApps: { color: '#5c2d91', label: 'Role-Based Apps' },
    system: { color: '#605e5c', label: 'System/Admin' },
  },

  /**
   * Core process icons (blue) - visible to all users
   */
  coreIcon: {
    color: '#0078d4',
    fontSize: '16px',
  },

  /**
   * Enterprise icons (teal) - visible to licensed users
   */
  enterpriseIcon: {
    color: '#03787C',
    fontSize: '16px',
  },

  /**
   * Role-based app icons (purple) - visible based on user role
   */
  roleIcon: {
    color: '#5c2d91',
    fontSize: '16px',
  },

  /**
   * System icons (gray) - admin only
   */
  systemIcon: {
    color: '#605e5c',
    fontSize: '16px',
  },

  /**
   * Quick link grid (app launcher style)
   */
  quickLinkGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '8px',
    padding: '16px',
  },

  /**
   * Quick link item
   */
  quickLinkItem: {
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center',
    gap: '8px',
    padding: '12px',
    borderRadius: '8px',
    cursor: 'pointer',
    textDecoration: 'none',
    transition: 'background-color 0.15s ease',
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
  },

  quickLinkIcon: {
    width: '40px',
    height: '40px',
    borderRadius: '8px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '20px',
    color: '#ffffff',
  },

  quickLinkLabel: {
    fontSize: '12px',
    color: '#323130',
    textAlign: 'center' as const,
    maxWidth: '70px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap' as const,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 39. SYSTEM ICONS STYLES (NAV-SYS)
// Source: User Section (notifications, help, settings, profile)
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlSystemIconStyles = {
  /**
   * System icon container (user section in header)
   */
  container: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },

  /**
   * Notification bell icon button
   */
  notificationButton: {
    position: 'relative' as const,
    width: '36px',
    height: '36px',
    borderRadius: '50%',
    backgroundColor: 'rgba(255, 255, 255, 0.1)',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'background 0.15s ease',
    ':hover': {
      backgroundColor: 'rgba(255, 255, 255, 0.2)',
    },
  },

  /**
   * Notification badge (red dot with count)
   */
  notificationBadge: {
    position: 'absolute' as const,
    top: '2px',
    right: '2px',
    minWidth: '16px',
    height: '16px',
    padding: '0 4px',
    backgroundColor: '#d13438',
    color: '#ffffff',
    borderRadius: '8px',
    fontSize: '10px',
    fontWeight: 600,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    border: '2px solid #0078d4',
  },

  /**
   * Help icon button
   */
  helpButton: {
    width: '36px',
    height: '36px',
    borderRadius: '50%',
    backgroundColor: 'rgba(255, 255, 255, 0.1)',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'background 0.15s ease',
    ':hover': {
      backgroundColor: 'rgba(255, 255, 255, 0.2)',
    },
  },

  /**
   * Settings icon button
   */
  settingsButton: {
    width: '36px',
    height: '36px',
    borderRadius: '50%',
    backgroundColor: 'rgba(255, 255, 255, 0.1)',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'background 0.15s ease',
    ':hover': {
      backgroundColor: 'rgba(255, 255, 255, 0.2)',
    },
  },

  /**
   * User profile avatar button
   */
  profileButton: {
    width: '32px',
    height: '32px',
    borderRadius: '50%',
    backgroundColor: '#8764b8',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '13px',
    fontWeight: 600,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  /**
   * Profile dropdown panel
   */
  profileDropdown: {
    position: 'absolute' as const,
    top: '100%',
    right: 0,
    marginTop: '8px',
    minWidth: '280px',
    backgroundColor: '#ffffff',
    borderRadius: '8px',
    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.15)',
    zIndex: 1000,
  },

  profileDropdownHeader: {
    padding: '16px',
    borderBottom: '1px solid #edebe9',
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },

  profileDropdownAvatar: {
    width: '48px',
    height: '48px',
    borderRadius: '50%',
    backgroundColor: '#8764b8',
    color: '#ffffff',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '18px',
    fontWeight: 600,
  },

  profileDropdownInfo: {
    flex: 1,
    minWidth: 0,
  },

  profileDropdownName: {
    fontSize: '14px',
    fontWeight: 600,
    color: '#323130',
    marginBottom: '2px',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  profileDropdownEmail: {
    fontSize: '12px',
    color: '#605e5c',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap' as const,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },

  profileDropdownItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '10px 16px',
    color: '#323130',
    textDecoration: 'none',
    fontSize: '14px',
    cursor: 'pointer',
    transition: 'background-color 0.1s ease',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      backgroundColor: '#f3f2f1',
    },
  },

  profileDropdownItemIcon: {
    color: '#0078d4',
    fontSize: '16px',
    width: '20px',
    textAlign: 'center' as const,
  },

  profileDropdownDivider: {
    height: '1px',
    backgroundColor: '#edebe9',
    margin: '8px 0',
  },

  profileDropdownFooter: {
    padding: '8px 0',
  },

  signOutButton: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '10px 16px',
    width: '100%',
    border: 'none',
    backgroundColor: 'transparent',
    color: '#d13438',
    fontSize: '14px',
    cursor: 'pointer',
    textAlign: 'left' as const,
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    ':hover': {
      backgroundColor: '#fde7e9',
    },
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// 40. SUBHEADER VARIANTS (NAV-SUB)
// Style 1: Blue Left Accent (Default) - Already in JmlSubheaderStyles
// Added: All 3 official sub-header variants
// ════════════════════════════════════════════════════════════════════════════════════

export const JmlSubheaderVariants = {
  /**
   * Style 1: Blue Left Accent (DEFAULT)
   * Source: .jml-subheader-accent
   */
  accentLeft: {
    container: {
      background: 'linear-gradient(135deg, #e8f4fd 0%, #f0f8ff 100%)',
      borderLeft: '4px solid #0078d4',
      borderRadius: '8px',
      padding: '16px 20px',
    },
    title: {
      fontSize: '20px',
      fontWeight: 600,
      color: '#004578',
      margin: '0 0 4px 0',
      fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    },
    description: {
      fontSize: '13px',
      color: '#605e5c',
      margin: 0,
      fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    },
  },

  /**
   * Style 2: Underline Accent with Step Badge
   * Source: .jml-subheader-underline
   */
  underline: {
    container: {
      backgroundColor: '#ffffff',
      borderBottom: '3px solid #0078d4',
      padding: '12px 16px',
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
    },
    stepBadge: {
      backgroundColor: '#0078d4',
      color: '#ffffff',
      width: '28px',
      height: '28px',
      borderRadius: '50%',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: '14px',
      fontWeight: 600,
      flexShrink: 0,
    },
    title: {
      fontSize: '18px',
      fontWeight: 600,
      color: '#323130',
      margin: 0,
      fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    },
  },

  /**
   * Style 3: Icon Banner (Modern)
   * Source: .jml-subheader-banner
   */
  iconBanner: {
    container: {
      backgroundColor: '#ffffff',
      border: '1px solid #e1e5e9',
      borderRadius: '12px',
      padding: '20px 24px',
      boxShadow: '0 2px 8px rgba(0, 0, 0, 0.06)',
      display: 'flex',
      alignItems: 'center',
      gap: '16px',
    },
    iconBox: {
      background: 'linear-gradient(135deg, #0078d4 0%, #004578 100%)',
      width: '52px',
      height: '52px',
      borderRadius: '12px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      color: '#ffffff',
      fontSize: '24px',
      flexShrink: 0,
    },
    content: {
      flex: 1,
      minWidth: 0,
    },
    title: {
      fontSize: '20px',
      fontWeight: 600,
      color: '#323130',
      margin: '0 0 4px 0',
      fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    },
    description: {
      fontSize: '13px',
      color: '#605e5c',
      margin: 0,
      fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    },
  },
};

// ============================================================================
// SECTION 41: FULL PAGE LAYOUT (LAY-FPD)
// Complete page structure for composing JML pages
// ============================================================================

/**
 * JML Full Page Layout Styles (LAY-FPD)
 * Complete page structure showing all components working together
 *
 * Usage:
 * <div style={JmlFullPageLayoutStyles.pageWrapper}>
 *   <header style={JmlFullPageLayoutStyles.header}>
 *     <!-- Nav bar + Page header -->
 *   </header>
 *   <main style={JmlFullPageLayoutStyles.mainContent}>
 *     <div style={JmlFullPageLayoutStyles.contentWrapper}>
 *       <!-- Page content -->
 *     </div>
 *   </main>
 *   <footer style={JmlFullPageLayoutStyles.footer}>
 *     <!-- Footer content -->
 *   </footer>
 * </div>
 */
export const JmlFullPageLayoutStyles = {
  /**
   * Page wrapper - full viewport container
   * Creates flex column layout for header/main/footer structure
   */
  pageWrapper: {
    display: 'flex',
    flexDirection: 'column' as const,
    minHeight: '100vh',
    backgroundColor: '#f3f2f1',
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', sans-serif",
  },

  /**
   * Header container - contains nav bar and page header
   * Gradient background with high z-index
   */
  header: {
    background: 'linear-gradient(135deg, #0078d4 0%, #004578 100%)',
    color: '#ffffff',
    position: 'relative' as const,
    zIndex: 100000,
    flexShrink: 0,
  },

  /**
   * Navigation bar within header
   * Height: 56px, horizontal flex layout
   */
  navBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '8px 24px',
    height: '56px',
    boxShadow: '0 2px 8px rgba(0, 0, 0, 0.2)',
  },

  /**
   * Page header (title section) within header
   * Darker gradient, contains breadcrumb + title
   */
  pageHeader: {
    background: 'linear-gradient(180deg, #004578 0%, #005a9e 100%)',
    padding: '16px 24px 20px',
    borderBottom: '4px solid #0078d4',
  },

  /**
   * Main content area - fills remaining space
   * White background with hidden overflow
   */
  mainContent: {
    flex: 1,
    backgroundColor: '#ffffff',
    overflowX: 'hidden' as const,
  },

  /**
   * Content wrapper - centers and constrains content
   * Max width 1400px with auto margins
   */
  contentWrapper: {
    maxWidth: '1400px',
    width: '100%',
    margin: '0 auto',
    padding: '24px',
  },

  /**
   * Full-width content wrapper variant
   * No max-width constraint, no padding
   */
  contentWrapperFullWidth: {
    maxWidth: 'none',
    width: '100%',
    margin: '0 auto',
    padding: 0,
  },

  /**
   * Demo/placeholder content area
   * For showing where content goes in mockups
   */
  demoContentArea: {
    backgroundColor: '#faf9f8',
    border: '2px dashed #edebe9',
    borderRadius: '12px',
    padding: '48px',
    textAlign: 'center' as const,
    color: '#a19f9d',
    minHeight: '200px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '14px',
    fontWeight: 500,
  },

  /**
   * Footer - fixed at bottom
   * Dark gradient matching header
   */
  footer: {
    backgroundColor: '#004578',
    color: 'rgba(255, 255, 255, 0.9)',
    flexShrink: 0,
  },

  /**
   * Z-index layers for proper stacking
   */
  zIndex: {
    header: 100000,
    navigation: 100001,
    dropdown: 100002,
    modal: 100003,
    tooltip: 100004,
  },

  /**
   * CSS Variables reference (for SCSS/CSS usage)
   */
  cssVariables: {
    maxWidth: '1400px',
    headerHeight: '56px',
    sidebarWidth: '280px',
    spacingXs: '4px',
    spacingS: '8px',
    spacingM: '12px',
    spacingL: '16px',
    spacingXl: '24px',
    spacingXxl: '32px',
    spacingXxxl: '48px',
  },
};

// ============================================================================
// SECTION 42: MAIN CONTENT CONTAINER (LAY-MCN)
// Content area patterns and variants
// ============================================================================

/**
 * JML Main Content Styles (LAY-MCN)
 * Patterns for main content area layout
 */
export const JmlMainContentStyles = {
  /**
   * Standard content container
   * Centered with max-width constraint
   */
  container: {
    maxWidth: '1400px',
    width: '100%',
    margin: '0 auto',
    padding: '24px',
    boxSizing: 'border-box' as const,
  },

  /**
   * Full-width container
   * No max-width, edge-to-edge content
   */
  containerFullWidth: {
    maxWidth: 'none',
    width: '100%',
    margin: 0,
    padding: 0,
    boxSizing: 'border-box' as const,
  },

  /**
   * Narrow container for focused content
   * 800px max-width, good for forms/wizards
   */
  containerNarrow: {
    maxWidth: '800px',
    width: '100%',
    margin: '0 auto',
    padding: '24px',
    boxSizing: 'border-box' as const,
  },

  /**
   * Wide container for dashboards
   * 1600px max-width
   */
  containerWide: {
    maxWidth: '1600px',
    width: '100%',
    margin: '0 auto',
    padding: '24px',
    boxSizing: 'border-box' as const,
  },

  /**
   * Content section spacing
   */
  section: {
    marginBottom: '24px',
  },

  /**
   * Content section with card style
   */
  sectionCard: {
    backgroundColor: '#ffffff',
    borderRadius: '8px',
    padding: '24px',
    boxShadow: '0 1px 4px rgba(0, 0, 0, 0.08)',
    marginBottom: '24px',
  },

  /**
   * Two-column layout
   */
  twoColumn: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '24px',
  },

  /**
   * Three-column layout
   */
  threeColumn: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '24px',
  },

  /**
   * Sidebar + content layout
   * 280px sidebar + flexible content
   */
  sidebarLayout: {
    display: 'grid',
    gridTemplateColumns: '280px 1fr',
    gap: '24px',
    minHeight: 'calc(100vh - 180px)',
  },

  /**
   * Sidebar container
   */
  sidebar: {
    backgroundColor: '#faf9f8',
    borderRadius: '8px',
    padding: '16px',
    height: 'fit-content',
    position: 'sticky' as const,
    top: '24px',
  },

  /**
   * Main content area in sidebar layout
   */
  mainArea: {
    minWidth: 0,
  },
};

// ============================================================================
// SECTION 43: OVERFLOW & TEXT HANDLING UTILITIES (REF-OVF)
// Patterns for preventing overflow issues
// ============================================================================

/**
 * JML Overflow Styles (REF-OVF)
 * Utility patterns for handling overflow and text truncation
 */
export const JmlOverflowStyles = {
  // =========================================================================
  // CONTAINER OVERFLOW
  // =========================================================================

  /**
   * overflow: hidden
   * Use for: Card containers (prevents border-radius bleed), modals, panels
   */
  hidden: {
    overflow: 'hidden' as const,
  },

  /**
   * overflow: auto
   * Use for: Panel content areas, modal bodies, scrollable containers
   */
  auto: {
    overflow: 'auto' as const,
  },

  /**
   * overflow-x: auto
   * Use for: Wide data tables, code blocks, horizontal scroll containers
   */
  horizontalScroll: {
    overflowX: 'auto' as const,
    overflowY: 'hidden' as const,
  },

  /**
   * overflow-y: auto
   * Use for: Tall content areas, side panels, long lists
   */
  verticalScroll: {
    overflowX: 'hidden' as const,
    overflowY: 'auto' as const,
  },

  /**
   * overflow: visible
   * Use for: Dropdown trigger containers, tooltip hosts, menus
   */
  visible: {
    overflow: 'visible' as const,
  },

  // =========================================================================
  // TEXT TRUNCATION
  // =========================================================================

  /**
   * Single line truncate with ellipsis
   * Use for: Table cells, nav links, card titles, list items
   */
  truncate: {
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden' as const,
    textOverflow: 'ellipsis',
  },

  /**
   * Two-line clamp with ellipsis
   * Use for: Card descriptions, preview text, comments
   */
  clamp2: {
    display: '-webkit-box',
    WebkitLineClamp: 2,
    WebkitBoxOrient: 'vertical' as const,
    overflow: 'hidden' as const,
  },

  /**
   * Three-line clamp with ellipsis
   */
  clamp3: {
    display: '-webkit-box',
    WebkitLineClamp: 3,
    WebkitBoxOrient: 'vertical' as const,
    overflow: 'hidden' as const,
  },

  /**
   * Four-line clamp with ellipsis
   */
  clamp4: {
    display: '-webkit-box',
    WebkitLineClamp: 4,
    WebkitBoxOrient: 'vertical' as const,
    overflow: 'hidden' as const,
  },

  // =========================================================================
  // SCROLLABLE CONTAINERS
  // =========================================================================

  /**
   * Scrollable panel content
   * Use for: Panel body, modal body, side panel content
   */
  scrollablePanel: {
    flex: 1,
    overflowY: 'auto' as const,
    overflowX: 'hidden' as const,
    padding: '24px',
  },

  /**
   * Scrollable table container
   * Use for: Wide tables that need horizontal scroll
   */
  scrollableTable: {
    overflowX: 'auto' as const,
    width: '100%',
    WebkitOverflowScrolling: 'touch',
  },

  /**
   * Scrollable list
   * Use for: Long lists, file lists, item lists
   */
  scrollableList: {
    maxHeight: '400px',
    overflowY: 'auto' as const,
    overflowX: 'hidden' as const,
  },

  // =========================================================================
  // OVERFLOW PREVENTION FOR FLEX/GRID
  // =========================================================================

  /**
   * Prevent flex child overflow
   * Use for: Flex children that might overflow parent
   */
  flexSafe: {
    minWidth: 0,
    overflow: 'hidden' as const,
  },

  /**
   * Prevent grid child overflow
   * Use for: Grid children that might overflow cell
   */
  gridSafe: {
    minWidth: 0,
    maxWidth: '100%',
    overflow: 'hidden' as const,
  },

  /**
   * Word break for long text
   * Use for: Long URLs, email addresses, unbroken strings
   */
  breakWord: {
    wordBreak: 'break-word' as const,
    overflowWrap: 'break-word' as const,
  },
};

// ════════════════════════════════════════════════════════════════════════════════════
// EDGE CASE PATTERNS - Contract Manager & Cross-Component Styles
// Added for complete consistency across all JML views
// ════════════════════════════════════════════════════════════════════════════════════

// ────────────────────────────────────────────────────────────────────────────────────
// 1. ALERT/NOTIFICATION LIST STYLES
// Use for: Dashboard alerts, notification lists, warning items
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlAlertListStyles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.s,
  },

  item: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
    padding: FluentSpacing.m,
    borderRadius: '6px',
    backgroundColor: FluentColors.neutralLighterAlt,
    cursor: 'pointer',
    transition: 'background-color 0.2s ease',
    selectors: {
      ':hover': {
        backgroundColor: FluentColors.neutralLight,
      },
    },
  },

  icon: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },

  content: {
    flex: 1,
    minWidth: 0,
  },

  title: {
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    fontSize: '14px',
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  subtitle: {
    color: FluentColors.neutralSecondary,
    fontSize: '12px',
    margin: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  // Alert severity variants
  severityColors: {
    critical: { background: '#fde7e9', iconColor: '#d13438' },
    warning: { background: '#fff4ce', iconColor: '#ffaa00' },
    info: { background: '#e8f4fd', iconColor: '#0078d4' },
    success: { background: '#dff6dd', iconColor: '#107c10' },
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 2. TIMELINE/ACTIVITY STYLES
// Use for: Activity history, audit logs, process timelines
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlTimelineStyles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
  },

  item: {
    display: 'flex',
    gap: FluentSpacing.m,
    padding: FluentSpacing.m,
    position: 'relative' as const,
  },

  itemWithLine: {
    display: 'flex',
    gap: FluentSpacing.m,
    padding: FluentSpacing.m,
    position: 'relative' as const,
    selectors: {
      ':not(:last-child)::before': {
        content: '""',
        position: 'absolute',
        left: '18px',
        top: '44px',
        bottom: '-8px',
        width: '2px',
        backgroundColor: FluentColors.neutralLight,
      },
    },
  },

  dot: {
    width: '12px',
    height: '12px',
    borderRadius: '50%',
    backgroundColor: FluentColors.themePrimary,
    flexShrink: 0,
    marginTop: '4px',
  },

  dotLarge: {
    width: '36px',
    height: '36px',
    borderRadius: '50%',
    backgroundColor: FluentColors.themeLighter,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
    color: FluentColors.themePrimary,
  },

  content: {
    flex: 1,
    minWidth: 0,
  },

  title: {
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    fontSize: '14px',
    margin: '0 0 4px 0',
    fontFamily: FluentTypography.fontFamily,
  },

  description: {
    color: FluentColors.neutralSecondary,
    fontSize: '13px',
    margin: '0 0 4px 0',
    fontFamily: FluentTypography.fontFamily,
  },

  date: {
    fontSize: '12px',
    color: FluentColors.neutralTertiary,
    fontFamily: FluentTypography.fontFamily,
  },

  // Status-specific dot colors
  dotColors: {
    created: FluentColors.themePrimary,
    updated: '#ffaa00',
    approved: '#107c10',
    rejected: '#d13438',
    pending: FluentColors.neutralTertiary,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 3. PRIORITY BAR/INDICATOR STYLES
// Use for: Approval cards, obligation items, task priority indicators
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlPriorityBarStyles = {
  bar: {
    width: '4px',
    borderRadius: '4px',
    alignSelf: 'stretch',
    flexShrink: 0,
  },

  barHorizontal: {
    height: '4px',
    width: '100%',
    borderRadius: '4px',
  },

  // Priority color mapping
  colors: {
    critical: '#d13438',
    high: '#ff8c00',
    medium: '#0078d4',
    low: '#107c10',
    none: FluentColors.neutralLight,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 4. INTERACTIVE LIST ITEM CARD STYLES
// Use for: Approval cards, obligation items, contract selection lists
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlListItemCardStyles = {
  card: {
    display: 'flex',
    padding: FluentSpacing.l,
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    gap: FluentSpacing.m,
    alignItems: 'flex-start',
    transition: 'all 0.2s ease',
    cursor: 'pointer',
    selectors: {
      ':hover': {
        backgroundColor: FluentColors.neutralLighterAlt,
        borderColor: FluentColors.neutralTertiary,
        boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
      },
    },
  },

  cardSelected: {
    borderColor: FluentColors.themePrimary,
    backgroundColor: FluentColors.themeLighterAlt,
    selectors: {
      ':hover': {
        borderColor: FluentColors.themePrimary,
        backgroundColor: FluentColors.themeLighter,
      },
    },
  },

  content: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.xs,
    minWidth: 0,
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: FluentSpacing.m,
  },

  titleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.s,
    flexWrap: 'wrap' as const,
  },

  title: {
    fontSize: '14px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    margin: 0,
    lineHeight: 1.3,
    fontFamily: FluentTypography.fontFamily,
  },

  meta: {
    display: 'flex',
    gap: FluentSpacing.m,
    color: FluentColors.neutralSecondary,
    fontSize: '13px',
    flexWrap: 'wrap' as const,
  },

  metaItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },

  actions: {
    display: 'flex',
    gap: FluentSpacing.s,
    flexShrink: 0,
  },

  arrow: {
    display: 'flex',
    alignItems: 'center',
    color: FluentColors.neutralTertiary,
    flexShrink: 0,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 5. DETAIL GRID STYLES (Panel/Dialog Content)
// Use for: Fly-in panel details, dialog content grids, detail views
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlDetailGridStyles = {
  section: {
    marginBottom: FluentSpacing.l,
  },

  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: FluentSpacing.m,
  },

  gridThreeColumn: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: FluentSpacing.m,
    '@media (max-width: 768px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
  },

  item: {
    display: 'flex',
    flexDirection: 'column' as const,
  },

  itemFullWidth: {
    display: 'flex',
    flexDirection: 'column' as const,
    gridColumn: '1 / -1',
  },

  label: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    marginBottom: '2px',
    fontWeight: 500,
    fontFamily: FluentTypography.fontFamily,
  },

  value: {
    fontSize: '14px',
    fontWeight: 500,
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },

  valueLarge: {
    fontSize: '18px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 6. UPCOMING DATE WIDGET STYLES
// Use for: Calendar date displays, deadline widgets, upcoming events
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlUpcomingDateStyles = {
  container: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: FluentSpacing.m,
    padding: `${FluentSpacing.s} 0`,
    borderBottom: `1px solid ${FluentColors.neutralLighter}`,
    cursor: 'pointer',
    selectors: {
      ':last-child': {
        borderBottom: 'none',
      },
      ':hover': {
        backgroundColor: FluentColors.neutralLighterAlt,
      },
    },
  },

  dateBox: {
    width: '48px',
    height: '48px',
    borderRadius: '8px',
    backgroundColor: FluentColors.themeLighterAlt,
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },

  dateBoxUrgent: {
    backgroundColor: '#fde7e9',
  },

  dateBoxWarning: {
    backgroundColor: '#fff4ce',
  },

  day: {
    fontSize: '18px',
    fontWeight: 700,
    color: FluentColors.themePrimary,
    lineHeight: 1,
    fontFamily: FluentTypography.fontFamily,
  },

  dayUrgent: {
    color: '#d13438',
  },

  dayWarning: {
    color: '#b37700',
  },

  month: {
    fontSize: '11px',
    color: FluentColors.themePrimary,
    textTransform: 'uppercase' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  content: {
    flex: 1,
    minWidth: 0,
  },

  title: {
    fontWeight: 500,
    marginBottom: '2px',
    fontSize: '14px',
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },

  subtitle: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 7. TYPE/CATEGORY BREAKDOWN WIDGET STYLES
// Use for: Category summaries, type distributions, breakdown lists
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlTypeBreakdownStyles = {
  container: {
    marginTop: FluentSpacing.m,
  },

  row: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: `${FluentSpacing.s} 0`,
    borderBottom: `1px solid ${FluentColors.neutralLighter}`,
    selectors: {
      ':last-child': {
        borderBottom: 'none',
      },
    },
  },

  label: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.s,
  },

  icon: {
    width: '8px',
    height: '8px',
    borderRadius: '50%',
    flexShrink: 0,
  },

  iconSquare: {
    width: '8px',
    height: '8px',
    borderRadius: '2px',
    flexShrink: 0,
  },

  text: {
    fontSize: '14px',
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },

  value: {
    fontSize: '14px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    fontVariantNumeric: 'tabular-nums' as const,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 8. PAGINATION STYLES
// Use for: Data grid pagination, list navigation, page controls
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlPaginationStyles = {
  container: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: FluentSpacing.m,
    borderTop: `1px solid ${FluentColors.neutralLight}`,
  },

  containerBottom: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: FluentSpacing.l,
    borderTop: `1px solid ${FluentColors.neutralLight}`,
    backgroundColor: FluentColors.white,
  },

  info: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    fontFamily: FluentTypography.fontFamily,
  },

  infoHighlight: {
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
  },

  controls: {
    display: 'flex',
    gap: FluentSpacing.s,
    alignItems: 'center',
  },

  pageNumber: {
    minWidth: '32px',
    height: '32px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: '4px',
    fontSize: '14px',
    cursor: 'pointer',
    selectors: {
      ':hover': {
        backgroundColor: FluentColors.neutralLighter,
      },
    },
  },

  pageNumberActive: {
    backgroundColor: FluentColors.themePrimary,
    color: FluentColors.white,
    selectors: {
      ':hover': {
        backgroundColor: FluentColors.themeDarkAlt,
      },
    },
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 9. CHART SECTION STYLES
// Use for: Analytics charts, dashboard graphs, data visualizations
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlChartSectionStyles = {
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))',
    gap: FluentSpacing.l,
  },

  gridTwoColumn: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: FluentSpacing.l,
    '@media (max-width: 1000px)': {
      gridTemplateColumns: '1fr',
    },
  },

  card: {
    backgroundColor: FluentColors.white,
    borderRadius: '8px',
    border: `1px solid ${FluentColors.neutralLight}`,
    padding: FluentSpacing.l,
    boxShadow: '0 1px 3px rgba(0, 0, 0, 0.04)',
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: FluentSpacing.l,
  },

  title: {
    margin: 0,
    fontSize: '14px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    fontFamily: FluentTypography.fontFamily,
  },

  subtitle: {
    margin: '4px 0 0 0',
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    fontFamily: FluentTypography.fontFamily,
  },

  actions: {
    display: 'flex',
    gap: FluentSpacing.s,
  },

  chartContainer: {
    minHeight: '200px',
    position: 'relative' as const,
  },

  // Bar chart styles
  barContainer: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.m,
  },

  barItem: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
  },

  barLabel: {
    width: '100px',
    fontSize: '13px',
    color: FluentColors.neutralSecondary,
    flexShrink: 0,
    fontFamily: FluentTypography.fontFamily,
  },

  barWrapper: {
    flex: 1,
    height: '24px',
    backgroundColor: FluentColors.neutralLighter,
    borderRadius: '4px',
    overflow: 'hidden' as const,
  },

  barFill: {
    height: '100%',
    backgroundColor: FluentColors.themePrimary,
    borderRadius: '4px',
    transition: 'width 0.3s ease',
  },

  barValue: {
    width: '60px',
    textAlign: 'right' as const,
    fontSize: '13px',
    fontWeight: 600,
    fontVariantNumeric: 'tabular-nums' as const,
    fontFamily: FluentTypography.fontFamily,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 10. VALUE HIGHLIGHT STYLES
// Use for: Large monetary values, key metrics, highlighted numbers
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlValueHighlightStyles = {
  large: {
    fontSize: '24px',
    fontWeight: 600,
    color: FluentColors.themePrimary,
    fontVariantNumeric: 'tabular-nums' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  medium: {
    fontSize: '20px',
    fontWeight: 600,
    color: FluentColors.themePrimary,
    fontVariantNumeric: 'tabular-nums' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  small: {
    fontSize: '16px',
    fontWeight: 600,
    color: FluentColors.themePrimary,
    fontVariantNumeric: 'tabular-nums' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  // Color variants
  success: {
    color: '#107c10',
  },

  warning: {
    color: '#ff8c00',
  },

  danger: {
    color: '#d13438',
  },

  neutral: {
    color: FluentColors.neutralPrimary,
  },

  // Currency formatting helper (use with large/medium/small)
  currency: {
    fontFeatureSettings: '"tnum"',
    letterSpacing: '0.5px',
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 11. STATUS PROGRESS BAR STYLES
// Use for: Labeled progress indicators, status breakdowns, completion trackers
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlStatusProgressStyles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.m,
    marginTop: FluentSpacing.m,
  },

  row: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
  },

  label: {
    fontSize: '14px',
    color: FluentColors.neutralSecondary,
    minWidth: '100px',
    fontWeight: 600,
    fontFamily: FluentTypography.fontFamily,
  },

  labelWide: {
    minWidth: '140px',
  },

  progressWrapper: {
    flex: 1,
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.s,
  },

  progressBar: {
    flex: 1,
    height: '8px',
    backgroundColor: FluentColors.neutralLighter,
    borderRadius: '4px',
    overflow: 'hidden' as const,
  },

  progressFill: {
    height: '100%',
    borderRadius: '4px',
    transition: 'width 0.3s ease',
  },

  value: {
    fontSize: '14px',
    fontWeight: 600,
    fontVariantNumeric: 'tabular-nums' as const,
    minWidth: '50px',
    textAlign: 'right' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  count: {
    fontSize: '12px',
    color: FluentColors.neutralSecondary,
    minWidth: '30px',
    textAlign: 'right' as const,
    fontFamily: FluentTypography.fontFamily,
  },

  // Status colors for progress fills
  statusColors: {
    success: '#107c10',
    warning: '#ffaa00',
    danger: '#d13438',
    info: FluentColors.themePrimary,
    neutral: FluentColors.neutralSecondary,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 12. MENU STYLES (Section Headers & Action Buttons)
// Use for: Context menus, dropdown sections, action menu buttons
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlMenuStyles = {
  sectionHeader: {
    padding: '6px 12px',
    fontSize: '11px',
    color: FluentColors.neutralSecondary,
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    fontWeight: 600,
    fontFamily: FluentTypography.fontFamily,
  },

  divider: {
    height: '1px',
    backgroundColor: FluentColors.neutralLight,
    margin: '4px 0',
  },

  actionButton: {
    padding: '8px 16px',
    borderRadius: '4px',
    border: `1px solid ${FluentColors.themePrimary}`,
    background: 'white',
    color: FluentColors.themePrimary,
    cursor: 'pointer',
    fontSize: '14px',
    fontWeight: 500,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    transition: 'all 0.15s',
    fontFamily: FluentTypography.fontFamily,
    selectors: {
      ':hover': {
        background: '#f0f6fc',
      },
    },
  },

  actionButtonDanger: {
    border: `1px solid #d13438`,
    color: '#d13438',
    selectors: {
      ':hover': {
        background: '#fde7e9',
      },
    },
  },

  menuItemDanger: {
    color: '#d13438',
    selectors: {
      ':hover': {
        backgroundColor: '#fde7e9',
      },
    },
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 13. PANEL LAYOUT STYLES
// Use for: Fly-in panels, side panels, detail panels
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlPanelLayoutStyles = {
  content: {
    padding: FluentSpacing.l,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.l,
  },

  contentCompact: {
    padding: FluentSpacing.m,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.m,
  },

  header: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
  },

  headerWithIcon: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.m,
    paddingBottom: FluentSpacing.m,
    borderBottom: `1px solid ${FluentColors.neutralLight}`,
  },

  footer: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: FluentSpacing.s,
    padding: FluentSpacing.l,
    borderTop: `1px solid ${FluentColors.neutralLight}`,
    backgroundColor: FluentColors.white,
  },

  footerSpaceBetween: {
    display: 'flex',
    justifyContent: 'space-between',
    gap: FluentSpacing.s,
    padding: FluentSpacing.l,
    borderTop: `1px solid ${FluentColors.neutralLight}`,
    backgroundColor: FluentColors.white,
  },

  section: {
    marginBottom: FluentSpacing.l,
  },

  sectionTitle: {
    fontSize: '14px',
    fontWeight: 600,
    color: FluentColors.neutralPrimary,
    marginBottom: FluentSpacing.m,
    fontFamily: FluentTypography.fontFamily,
  },

  // For scrollable panel content
  scrollableContent: {
    flex: 1,
    overflowY: 'auto' as const,
    padding: FluentSpacing.l,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 14. CHANGE DIFF STYLES
// Use for: Audit logs, change tracking, version comparison
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlChangeDiffStyles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: FluentSpacing.xs,
  },

  row: {
    display: 'flex',
    gap: FluentSpacing.s,
    fontSize: '13px',
  },

  old: {
    color: '#d13438',
    textDecoration: 'line-through' as const,
  },

  new: {
    color: '#107c10',
  },

  arrow: {
    color: FluentColors.neutralTertiary,
  },

  field: {
    fontWeight: 600,
    color: FluentColors.neutralSecondary,
    marginRight: FluentSpacing.xs,
  },
};

// ────────────────────────────────────────────────────────────────────────────────────
// 15. FILTER CHIP STYLES
// Use for: Active filter display, tag lists, removable badges
// ────────────────────────────────────────────────────────────────────────────────────

export const JmlFilterChipStyles = {
  container: {
    display: 'flex',
    gap: FluentSpacing.s,
    flexWrap: 'wrap' as const,
  },

  chip: {
    display: 'flex',
    alignItems: 'center',
    gap: FluentSpacing.xs,
    padding: `${FluentSpacing.xs} ${FluentSpacing.s}`,
    backgroundColor: FluentColors.themeLighter,
    borderRadius: '16px',
    fontSize: '12px',
    color: FluentColors.themeDark,
    fontFamily: FluentTypography.fontFamily,
  },

  chipOutline: {
    backgroundColor: 'transparent',
    border: `1px solid ${FluentColors.themePrimary}`,
    color: FluentColors.themePrimary,
  },

  removeButton: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '16px',
    height: '16px',
    borderRadius: '50%',
    cursor: 'pointer',
    selectors: {
      ':hover': {
        backgroundColor: FluentColors.themeTertiary,
      },
    },
  },
};

export const JmlStyles = {
  // Layout & Structure
  pageHeader: JmlPageHeaderStyles,
  subheader: JmlSubheaderStyles,
  subheaderVariants: JmlSubheaderVariants,
  view: JmlViewStyles,
  section: JmlSectionStyles,
  grid: JmlGridStyles,
  fullPageLayout: JmlFullPageLayoutStyles,
  mainContent: JmlMainContentStyles,

  // Navigation
  navigation: JmlNavigationStyles,
  breadcrumb: JmlBreadcrumbStyles,
  tabPanel: JmlTabPanelStyles,
  navIcons: JmlNavIconStyles,
  systemIcons: JmlSystemIconStyles,
  footer: JmlFooterStyles,

  // Components
  stats: JmlStatsRowStyles,
  table: JmlTableStyles,
  tableResizable: JmlTableResizableStyles,
  commandPanel: JmlCommandPanelStyles,
  card: JmlCardStyles,

  // States
  empty: JmlEmptyStateStyles,
  loading: JmlLoadingStyles,
  filter: JmlFilterStyles,

  // Design Tokens
  colors: JmlColorPalette,
  typography: JmlTypographyStyles,
  spacing: JmlSpacingScale,
  borderRadius: JmlBorderRadius,
  shadows: JmlShadows,

  // Interactive Elements
  button: JmlButtonStyles,
  badge: JmlBadgeStyles,
  dropdown: JmlDropdownStyles,
  link: JmlLinkStyles,
  tooltip: JmlTooltipStyles,
  checkbox: JmlCheckboxStyles,
  radio: JmlRadioStyles,
  toggle: JmlToggleStyles,
  slider: JmlSliderStyles,

  // User/Identity
  avatar: JmlAvatarStyles,

  // Overlays
  panel: JmlPanelStyles,
  modal: JmlModalStyles,
  messageBar: JmlMessageBarStyles,

  // Forms
  form: JmlFormStyles,

  // Progress
  progress: JmlProgressStyles,

  // Layout Helpers
  divider: JmlDividerStyles,
  overflow: JmlOverflowStyles,

  // Edge Case Patterns (Contract Manager & Cross-Component)
  alertList: JmlAlertListStyles,
  timeline: JmlTimelineStyles,
  priorityBar: JmlPriorityBarStyles,
  listItemCard: JmlListItemCardStyles,
  detailGrid: JmlDetailGridStyles,
  upcomingDate: JmlUpcomingDateStyles,
  typeBreakdown: JmlTypeBreakdownStyles,
  pagination: JmlPaginationStyles,
  chartSection: JmlChartSectionStyles,
  valueHighlight: JmlValueHighlightStyles,
  statusProgress: JmlStatusProgressStyles,
  menu: JmlMenuStyles,
  panelLayout: JmlPanelLayoutStyles,
  changeDiff: JmlChangeDiffStyles,
  filterChip: JmlFilterChipStyles,
};

export default JmlStyles;
