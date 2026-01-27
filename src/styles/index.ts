/**
 * JML Styles - Central Export
 *
 * Import all JML styles from this single location:
 *
 * @example
 * import { JmlViewStyles, JmlStatsRowStyles, FluentColors } from '../../../../styles';
 *
 * Or import the unified namespace:
 * import { JmlStyles, Fluent } from '../../../../styles';
 */

// Fluent primitives and helpers
export {
  FluentColors,
  FluentTypography,
  FluentSpacing,
  FluentBorderRadius,
  FluentShadows,
  FluentAnimations,
  IconSizes,
  IconColors,
  getBadgeStyle,
  getButtonStyle,
  getCardStyle,
  getProgressBarStyle,
  getStandardHeaderStyles,
  getActionButtonStyle,
  getFilterRowStyles,
  getStatsCardGridStyles,
  getFullWidthContainerStyles,
  getFullWidthTableStyles,
  getStandardTableStyles,
  getStatCardStyle,
  getPageContainerStyles,
} from './FluentUIStyles';

// Tab panel styles (Fluent v9)
export { useTabPanelStyles, createTabDefinitions } from './TabPanelStyles';
export type { ITabDefinition } from './TabPanelStyles';

// Fluent v8 component styles
export {
  fluentColors,
  fluentShadows,
  fluentDropdownStyles,
  fluentDropdownCompactStyles,
  fluentDialogStyles,
  fluentPanelStyles,
  fluentCalloutStyles,
  fluentTooltipStyles,
  fluentTooltipInvertedStyles,
  fluentButtonStyles,
  getDropdownStyles,
  getDropdownCompactStyles,
} from './fluentV8Styles';

// JML View Layout Patterns (COMPOSED STYLES)
export {
  // Layout & Structure
  JmlPageHeaderStyles,
  JmlSubheaderStyles,
  JmlSubheaderVariants,
  JmlViewStyles,
  JmlSectionStyles,
  JmlGridStyles,
  JmlFullPageLayoutStyles,
  JmlMainContentStyles,

  // Navigation
  JmlNavigationStyles,
  JmlBreadcrumbStyles,
  JmlTabPanelStyles,
  JmlNavIconStyles,
  JmlSystemIconStyles,
  JmlFooterStyles,

  // Components
  JmlStatsRowStyles,
  JmlTableStyles,
  JmlTableResizableStyles,
  JmlCommandPanelStyles,
  JmlCardStyles,

  // States
  JmlEmptyStateStyles,
  JmlLoadingStyles,
  JmlFilterStyles,

  // Design Tokens
  JmlColorPalette,
  JmlTypographyStyles,
  JmlSpacingScale,
  JmlBorderRadius,
  JmlShadows,

  // Interactive Elements
  JmlButtonStyles,
  JmlBadgeStyles,
  JmlDropdownStyles,
  JmlLinkStyles,
  JmlSliderStyles,

  // Form Controls
  JmlFormStyles,
  JmlCheckboxStyles,
  JmlRadioStyles,
  JmlToggleStyles,

  // Progress & Loading
  JmlProgressStyles,

  // Overlays & Tooltips
  JmlPanelStyles,
  JmlModalStyles,
  JmlMessageBarStyles,
  JmlTooltipStyles,

  // Personas & Avatars
  JmlAvatarStyles,

  // Layout Utilities
  JmlDividerStyles,
  JmlOverflowStyles,

  // Edge Case Patterns (Contract Manager & Cross-Component)
  JmlAlertListStyles,
  JmlTimelineStyles,
  JmlPriorityBarStyles,
  JmlListItemCardStyles,
  JmlDetailGridStyles,
  JmlUpcomingDateStyles,
  JmlTypeBreakdownStyles,
  JmlPaginationStyles,
  JmlChartSectionStyles,
  JmlValueHighlightStyles,
  JmlStatusProgressStyles,
  JmlMenuStyles,
  JmlPanelLayoutStyles,
  JmlChangeDiffStyles,
  JmlFilterChipStyles,

  // Unified namespace
  JmlStyles,
} from './JmlViewStyles';

// Unified namespaces for convenience
export const Fluent = {
  Colors: {} as typeof import('./FluentUIStyles').FluentColors,
  Typography: {} as typeof import('./FluentUIStyles').FluentTypography,
  Spacing: {} as typeof import('./FluentUIStyles').FluentSpacing,
  BorderRadius: {} as typeof import('./FluentUIStyles').FluentBorderRadius,
  Shadows: {} as typeof import('./FluentUIStyles').FluentShadows,
  Animations: {} as typeof import('./FluentUIStyles').FluentAnimations,
};

// Initialize the unified namespace
import * as FluentUI from './FluentUIStyles';
Fluent.Colors = FluentUI.FluentColors;
Fluent.Typography = FluentUI.FluentTypography;
Fluent.Spacing = FluentUI.FluentSpacing;
Fluent.BorderRadius = FluentUI.FluentBorderRadius;
Fluent.Shadows = FluentUI.FluentShadows;
Fluent.Animations = FluentUI.FluentAnimations;
