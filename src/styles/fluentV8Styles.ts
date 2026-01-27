/**
 * Fluent UI v8 Consistent Styles - "Elevated & Clean" Design
 *
 * JML Design Standard for v8 components (Dropdown, Dialog, Panel)
 * Provides consistent styling that matches the JML solution aesthetic.
 *
 * Style: Elevated & Clean
 * - Balanced 4-8px border radius
 * - Subtle shadows with borders
 * - Left accent border on selected items
 * - Clean, professional appearance
 *
 * @see docs/JML-Design-System/dropdown-modal-styles.md
 */

import { IDropdownStyles, IDialogStyles, IPanelStyles, ICalloutContentStyles } from '@fluentui/react';

/**
 * Color palette for consistent theming
 */
export const fluentColors = {
  // Primary
  themePrimary: '#0078d4',
  themeDark: '#106ebe',
  themeLight: '#c7e0f4',

  // Neutrals
  white: '#ffffff',
  neutralLighter: '#f3f2f1',
  neutralLight: '#edebe9',
  neutralQuaternary: '#d1d1d1',
  neutralTertiary: '#a19f9d',
  neutralSecondary: '#605e5c',
  neutralPrimary: '#323130',

  // Borders (Elevated & Clean specific)
  borderLight: '#e8e8e8',
  borderMedium: '#e0e0e0',
  borderDark: '#d0d0d0',

  // Status
  success: '#107c10',
  warning: '#ffb900',
  error: '#a4262c',
  info: '#0078d4',

  // Selected states
  selectedBackground: '#e5f1fb',
  selectedBackgroundHover: '#cce4f7',
  hoverBackground: '#f3f2f1',

  // Disabled
  disabledBackground: '#f3f2f1',
  disabledText: '#a6a6a6'
};

/**
 * Shadow definitions
 */
export const fluentShadows = {
  // Dropdown/callout shadow
  elevation4: '0 1px 2px rgba(0, 0, 0, 0.08)',
  elevation8: '0 2px 4px rgba(0, 0, 0, 0.12)',
  elevation16: '0 6.4px 14.4px 0 rgba(0, 0, 0, 0.132), 0 1.2px 3.6px 0 rgba(0, 0, 0, 0.108)',
  // Dialog shadow
  elevation64: '0 25.6px 57.6px 0 rgba(0, 0, 0, 0.22), 0 4.8px 14.4px 0 rgba(0, 0, 0, 0.18)',
  // Primary button shadow
  primaryButton: '0 2px 4px rgba(0, 120, 212, 0.3)'
};

/**
 * Fluent "Elevated & Clean" dropdown styles for v8 Dropdown component
 */
export const fluentDropdownStyles: Partial<IDropdownStyles> = {
  root: {
    width: '100%'
  },
  dropdown: {
    borderRadius: '4px',
    border: `1px solid ${fluentColors.borderMedium}`,
    boxShadow: fluentShadows.elevation4,
    selectors: {
      ':hover': {
        borderColor: fluentColors.borderDark,
        boxShadow: fluentShadows.elevation8
      },
      ':focus': {
        borderColor: fluentColors.themePrimary,
        boxShadow: `0 0 0 2px rgba(0, 120, 212, 0.3)`
      },
      ':after': {
        borderRadius: '4px',
        border: `2px solid ${fluentColors.themePrimary}`
      }
    }
  },
  title: {
    borderRadius: '4px',
    border: 'none',
    backgroundColor: fluentColors.white,
    color: fluentColors.neutralPrimary,
    fontSize: '14px',
    fontWeight: 400,
    padding: '0 32px 0 12px',
    height: '32px',
    lineHeight: '30px'
  },
  caretDownWrapper: {
    right: '10px',
    top: '0',
    height: '32px',
    lineHeight: '32px'
  },
  caretDown: {
    color: fluentColors.neutralSecondary,
    fontSize: '12px'
  },
  dropdownItemsWrapper: {
    backgroundColor: fluentColors.white,
    borderRadius: '4px',
    boxShadow: fluentShadows.elevation16,
    border: `1px solid ${fluentColors.borderLight}`
  },
  dropdownItems: {
    padding: '4px 0'
  },
  dropdownItem: {
    backgroundColor: fluentColors.white,
    color: fluentColors.neutralPrimary,
    fontSize: '14px',
    padding: '8px 12px',
    minHeight: '32px',
    selectors: {
      ':hover': {
        backgroundColor: fluentColors.hoverBackground,
        color: fluentColors.neutralPrimary
      }
    }
  },
  dropdownItemSelected: {
    backgroundColor: fluentColors.selectedBackground,
    color: fluentColors.themePrimary,
    fontWeight: 600,
    borderLeft: `3px solid ${fluentColors.themePrimary}`,
    paddingLeft: '9px',
    selectors: {
      ':hover': {
        backgroundColor: fluentColors.selectedBackgroundHover,
        color: fluentColors.themePrimary
      }
    }
  },
  dropdownItemDisabled: {
    backgroundColor: fluentColors.disabledBackground,
    color: fluentColors.disabledText
  },
  callout: {
    borderRadius: '4px',
    boxShadow: fluentShadows.elevation16,
    border: `1px solid ${fluentColors.borderLight}`
  },
  errorMessage: {
    color: fluentColors.error,
    fontSize: '12px',
    marginTop: '4px'
  },
  label: {
    fontSize: '14px',
    fontWeight: 600,
    color: fluentColors.neutralPrimary,
    marginBottom: '4px'
  }
};

/**
 * Compact version for inline/toolbar dropdowns
 */
export const fluentDropdownCompactStyles: Partial<IDropdownStyles> = {
  root: {
    width: '100%'
  },
  dropdown: {
    borderRadius: '4px',
    border: `1px solid ${fluentColors.borderMedium}`,
    boxShadow: fluentShadows.elevation4,
    selectors: {
      ':hover': {
        borderColor: fluentColors.borderDark,
        boxShadow: fluentShadows.elevation8
      },
      ':focus': {
        borderColor: fluentColors.themePrimary,
        boxShadow: `0 0 0 2px rgba(0, 120, 212, 0.3)`
      },
      ':after': {
        borderRadius: '4px',
        border: `2px solid ${fluentColors.themePrimary}`
      }
    }
  },
  title: {
    borderRadius: '4px',
    border: 'none',
    backgroundColor: fluentColors.white,
    color: fluentColors.neutralPrimary,
    height: '28px',
    lineHeight: '26px',
    padding: '0 28px 0 8px',
    fontSize: '13px'
  },
  caretDownWrapper: {
    right: '6px',
    height: '28px',
    lineHeight: '28px'
  },
  caretDown: {
    color: fluentColors.neutralSecondary,
    fontSize: '12px'
  },
  dropdownItemsWrapper: {
    backgroundColor: fluentColors.white,
    borderRadius: '4px',
    boxShadow: fluentShadows.elevation16,
    border: `1px solid ${fluentColors.borderLight}`
  },
  dropdownItems: {
    padding: '4px 0'
  },
  dropdownItem: {
    backgroundColor: fluentColors.white,
    color: fluentColors.neutralPrimary,
    padding: '6px 10px',
    minHeight: '28px',
    fontSize: '13px',
    selectors: {
      ':hover': {
        backgroundColor: fluentColors.hoverBackground,
        color: fluentColors.neutralPrimary
      }
    }
  },
  dropdownItemSelected: {
    backgroundColor: fluentColors.selectedBackground,
    color: fluentColors.themePrimary,
    fontWeight: 600,
    borderLeft: `3px solid ${fluentColors.themePrimary}`,
    paddingLeft: '7px',
    selectors: {
      ':hover': {
        backgroundColor: fluentColors.selectedBackgroundHover,
        color: fluentColors.themePrimary
      }
    }
  },
  dropdownItemDisabled: {
    backgroundColor: fluentColors.disabledBackground,
    color: fluentColors.disabledText
  },
  callout: {
    borderRadius: '4px',
    boxShadow: fluentShadows.elevation16,
    border: `1px solid ${fluentColors.borderLight}`
  },
  errorMessage: {
    color: fluentColors.error,
    fontSize: '12px',
    marginTop: '4px'
  },
  label: {
    fontSize: '14px',
    fontWeight: 600,
    color: fluentColors.neutralPrimary,
    marginBottom: '4px'
  }
};

/**
 * Fluent "Elevated & Clean" dialog styles for v8 Dialog component
 */
export const fluentDialogStyles: Partial<IDialogStyles> = {
  main: {
    backgroundColor: fluentColors.white,
    borderRadius: '8px',
    boxShadow: fluentShadows.elevation64,
    border: `1px solid ${fluentColors.borderLight}`,
    padding: '24px',
    maxWidth: '600px',
    minWidth: '400px'
  }
};

/**
 * Fluent "Elevated & Clean" panel styles for v8 Panel component
 */
export const fluentPanelStyles: Partial<IPanelStyles> = {
  main: {
    backgroundColor: fluentColors.white,
    boxShadow: '-6.4px 0 14.4px 0 rgba(0, 0, 0, 0.132)'
  },
  content: {
    padding: '24px'
  },
  header: {
    padding: '16px 24px',
    borderBottom: `1px solid ${fluentColors.neutralLight}`
  },
  headerText: {
    fontSize: '20px',
    fontWeight: 600,
    color: fluentColors.neutralPrimary
  },
  footer: {
    padding: '16px 24px',
    borderTop: `1px solid ${fluentColors.neutralLight}`
  },
  footerInner: {
    display: 'flex',
    gap: '8px',
    justifyContent: 'flex-end'
  },
  commands: {
    marginTop: 0
  },
  navigation: {
    display: 'flex',
    justifyContent: 'flex-end'
  }
};

/**
 * Callout styles for consistent popover appearance
 */
export const fluentCalloutStyles: Partial<ICalloutContentStyles> = {
  root: {
    borderRadius: '8px',
    boxShadow: fluentShadows.elevation16,
    border: `1px solid ${fluentColors.neutralQuaternary}`
  },
  calloutMain: {
    backgroundColor: fluentColors.white,
    borderRadius: '8px',
    padding: '16px'
  },
  container: {
    backgroundColor: fluentColors.white
  },
  beak: {
    backgroundColor: fluentColors.white
  },
  beakCurtain: {
    backgroundColor: fluentColors.white
  }
};

/**
 * Tooltip styles
 */
export const fluentTooltipStyles = {
  root: {
    backgroundColor: fluentColors.neutralPrimary,
    color: fluentColors.white,
    borderRadius: '4px',
    padding: '8px 12px',
    fontSize: '12px',
    boxShadow: fluentShadows.elevation16
  }
};

/**
 * Inverted tooltip styles (light background)
 */
export const fluentTooltipInvertedStyles = {
  root: {
    backgroundColor: fluentColors.white,
    color: fluentColors.neutralPrimary,
    borderRadius: '4px',
    padding: '8px 12px',
    fontSize: '12px',
    border: `1px solid ${fluentColors.neutralQuaternary}`,
    boxShadow: fluentShadows.elevation16
  }
};

/**
 * Helper function to get dropdown styles with custom overrides
 */
export function getDropdownStyles(customStyles?: Partial<IDropdownStyles>): Partial<IDropdownStyles> {
  if (!customStyles) {
    return fluentDropdownStyles;
  }

  // Merge custom styles with defaults
  const merged: Partial<IDropdownStyles> = { ...fluentDropdownStyles };

  if (customStyles.root) {
    merged.root = { ...(merged.root as object), ...(customStyles.root as object) };
  }
  if (customStyles.dropdown) {
    merged.dropdown = { ...(merged.dropdown as object), ...(customStyles.dropdown as object) };
  }
  if (customStyles.title) {
    merged.title = { ...(merged.title as object), ...(customStyles.title as object) };
  }

  return merged;
}

/**
 * Helper function to get compact dropdown styles with custom overrides
 */
export function getDropdownCompactStyles(customStyles?: Partial<IDropdownStyles>): Partial<IDropdownStyles> {
  if (!customStyles) {
    return fluentDropdownCompactStyles;
  }

  // Merge custom styles with defaults
  const merged: Partial<IDropdownStyles> = { ...fluentDropdownCompactStyles };

  if (customStyles.root) {
    merged.root = { ...(merged.root as object), ...(customStyles.root as object) };
  }
  if (customStyles.dropdown) {
    merged.dropdown = { ...(merged.dropdown as object), ...(customStyles.dropdown as object) };
  }
  if (customStyles.title) {
    merged.title = { ...(merged.title as object), ...(customStyles.title as object) };
  }

  return merged;
}

/**
 * Button styles for use in dialogs and panels
 */
export const fluentButtonStyles = {
  primary: {
    root: {
      backgroundColor: fluentColors.themePrimary,
      color: fluentColors.white,
      borderRadius: '4px',
      border: 'none',
      boxShadow: fluentShadows.primaryButton
    },
    rootHovered: {
      backgroundColor: fluentColors.themeDark
    }
  },
  secondary: {
    root: {
      backgroundColor: fluentColors.white,
      color: fluentColors.neutralPrimary,
      borderRadius: '4px',
      border: `1px solid ${fluentColors.neutralTertiary}`
    },
    rootHovered: {
      backgroundColor: fluentColors.hoverBackground
    }
  }
};
