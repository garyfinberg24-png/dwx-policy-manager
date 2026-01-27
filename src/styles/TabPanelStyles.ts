/**
 * JML Standard Tab Panel Styles - Pill/Filled Style
 *
 * Consistent tab navigation styling following JML Design System
 * Features: White card background, pill-style active tab with blue fill
 *
 * Usage: Import useTabPanelStyles in your component
 *
 * @example
 * import { useTabPanelStyles } from '../styles/TabPanelStyles';
 * const tabStyles = useTabPanelStyles();
 *
 * <div className={tabStyles.tabPanel}>
 *   <div className={tabStyles.tabList} role="tablist">
 *     {tabs.map(tab => (
 *       <button
 *         role="tab"
 *         aria-selected={selectedTab === tab.value}
 *         className={mergeClasses(
 *           tabStyles.tab,
 *           selectedTab === tab.value && tabStyles.tabActive
 *         )}
 *         onClick={() => setSelectedTab(tab.value)}
 *       >
 *         {tab.icon}
 *         <span>{tab.label}</span>
 *       </button>
 *     ))}
 *   </div>
 *   <div className={tabStyles.panelActions}>
 *     // Optional action buttons
 *   </div>
 * </div>
 */

import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

/**
 * Primary Tab Panel Styles - JML Standard Pill/Filled Style
 * White background card with pill-style active tabs (blue fill)
 */
export const useTabPanelStyles = makeStyles({
  // ============================================
  // TAB PANEL CONTAINER - White Card Style
  // With blue accent bar on left border
  // ============================================
  tabPanel: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: tokens.colorNeutralBackground1,
    // Margins on all sides - matches Finance Manager pattern
    // Tab panel owns its inset from container edges
    marginTop: '12px',
    marginBottom: '20px', // Gap between tab panel and content below
    marginLeft: '24px',
    marginRight: '24px',
    // Blue accent on left border with 8px border radius on all corners
    borderLeft: '4px solid #0078d4',
    borderTopLeftRadius: '8px',
    borderBottomLeftRadius: '8px',
    borderTopRightRadius: '8px',
    borderBottomRightRadius: '8px',
    boxShadow: tokens.shadow4,
    // Padding inside the card
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    // Ensure proper stacking
    position: 'relative',
    zIndex: 1
  },

  // ============================================
  // TAB LIST - Container for tab buttons
  // ============================================
  tabList: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS
  },

  // ============================================
  // TAB BUTTON - Individual tab (inactive state)
  // ============================================
  tab: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    paddingTop: '10px',
    paddingBottom: '10px',
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground2,
    backgroundColor: 'transparent',
    ...shorthands.border('0'),
    ...shorthands.borderRadius('4px'),
    cursor: 'pointer',
    transitionProperty: 'all',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    whiteSpace: 'nowrap',
    ':hover': {
      color: tokens.colorBrandForeground1,
      backgroundColor: tokens.colorNeutralBackground1Hover
    },
    ':focus-visible': {
      ...shorthands.outline('2px', 'solid', tokens.colorBrandStroke1),
      outlineOffset: '2px'
    }
  },

  // ============================================
  // ACTIVE TAB - Pill/Filled Style (Blue Background)
  // ============================================
  tabActive: {
    color: '#ffffff',
    fontWeight: tokens.fontWeightSemibold,
    backgroundColor: '#0078d4',
    ':hover': {
      color: '#ffffff',
      backgroundColor: '#106ebe'
    }
  },

  // ============================================
  // TAB ICON
  // ============================================
  tabIcon: {
    fontSize: '16px',
    display: 'flex',
    alignItems: 'center'
  },

  // ============================================
  // PANEL ACTIONS - Right side buttons
  // ============================================
  panelActions: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center'
  },

  // ============================================
  // ACTION BUTTON - Icon button style
  // ============================================
  actionButton: {
    width: '36px',
    height: '36px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    ...shorthands.border('0'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    cursor: 'pointer',
    transitionProperty: 'all',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3Hover,
      color: tokens.colorNeutralForeground1
    },
    ':focus-visible': {
      ...shorthands.outline('2px', 'solid', tokens.colorBrandStroke1),
      outlineOffset: '2px'
    }
  },

  // ============================================
  // TAB CONTENT - Below tab panel
  // ============================================
  tabContent: {
    marginTop: tokens.spacingVerticalL,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL
  },

  tabContentNoPadding: {
    marginTop: tokens.spacingVerticalL
  },

  // ============================================
  // TAB PANEL FULL WIDTH - For use when container already has padding
  // With blue accent bar on left border
  // ============================================
  tabPanelFullWidth: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: tokens.colorNeutralBackground1,
    // Margins on all sides - matches Finance Manager pattern
    marginTop: '12px',
    marginBottom: '20px',
    marginLeft: '24px',
    marginRight: '24px',
    // Blue accent on left border with 8px border radius on all corners
    borderLeft: '4px solid #0078d4',
    borderTopLeftRadius: '8px',
    borderBottomLeftRadius: '8px',
    borderTopRightRadius: '8px',
    borderBottomRightRadius: '8px',
    boxShadow: tokens.shadow4,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    position: 'relative',
    zIndex: 1
  },

  // ============================================
  // TAB SUBHEADER - Style Option A (Blue Accent Bar)
  // For wizard step indicators and section titles
  // Aligned with tab panel margins (24px from edges)
  // ============================================
  tabSubheader: {
    backgroundColor: '#f0f6fc',
    backgroundImage: 'linear-gradient(135deg, #f0f6fc 0%, #e8f4fd 100%)',
    borderLeft: '4px solid #0078d4',
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL),
    // Full width to match tab panel
    marginLeft: 0,
    marginRight: 0,
    marginTop: tokens.spacingVerticalM,
    ...shorthands.borderRadius('8px'),
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between'
  },

  tabSubheaderContent: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('4px')
  },

  tabSubheaderTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: '#0078d4',
    lineHeight: '1.3'
  },

  tabSubheaderSubtitle: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2
  },

  // ============================================
  // WIZARD FOOTER - Navigation buttons
  // ============================================
  wizardFooter: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL),
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.borderRadius('0', '0', '8px', '8px'),
    marginTop: 'auto'
  },

  wizardFooterLeft: {
    display: 'flex',
    alignItems: 'center'
  },

  wizardFooterCenter: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2
  },

  wizardFooterRight: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS)
  }
});

/**
 * Tab definition interface for consistent tab structure
 */
export interface ITabDefinition<T extends string = string> {
  value: T;
  label: string;
  icon?: React.ReactNode;
  disabled?: boolean;
  badge?: number | string;
}

/**
 * Helper to create tab definitions
 */
export function createTabDefinitions<T extends string>(
  tabs: Array<{ value: T; label: string; icon?: React.ReactNode; disabled?: boolean; badge?: number | string }>
): ITabDefinition<T>[] {
  return tabs;
}
