// @ts-nocheck
import * as React from 'react';
import { Panel, IPanelProps, IPanelStyles } from '@fluentui/react';

/**
 * StyledPanel — standardised Fluent UI Panel wrapper for Policy Manager.
 *
 * Applies the Forest Teal panel styling consistently:
 * - Header: light teal gradient with teal border
 * - Close button: teal colored with hover effect
 * - Controls: 4px border-radius (inherited from global styles)
 * - Footer: light gray background with top border
 * - Surface: no border-radius (flush to edge)
 *
 * Usage: Drop-in replacement for <Panel>. All IPanelProps are passed through.
 *
 * EXCEPTION: The Policy Acknowledgement panel in PolicyDetails has its own
 * specific styling and does NOT use this wrapper.
 */

const PANEL_STYLES: Partial<IPanelStyles> = {
  main: {
    borderRadius: 0,
  },
  navigation: {
    background: 'var(--pm-panel-header-bg, linear-gradient(135deg, #f0fdfa 0%, #ccfbf1 100%))',
    height: 'auto',
    borderBottom: 'none',
    justifyContent: 'flex-end',
    padding: '8px 8px 0 0',
  },
  commands: {
    background: 'var(--pm-panel-header-bg, linear-gradient(135deg, #f0fdfa 0%, #ccfbf1 100%))',
    margin: 0,
    padding: '4px 4px 0 0',
  },
  header: {
    background: 'var(--pm-panel-header-bg, linear-gradient(135deg, #f0fdfa 0%, #ccfbf1 100%))',
    padding: '0 24px 12px 24px',
    marginBottom: 0,
    borderBottom: '1px solid var(--pm-primary-light, #99f6e4)',
  },
  headerText: {
    color: 'var(--pm-primary-dark, #0f766e)',
    fontWeight: 700,
    fontSize: '18px',
    lineHeight: '24px',
  },
  closeButton: {
    borderRadius: '4px',
    color: 'var(--pm-primary-dark, #0f766e)',
    selectors: {
      ':hover': {
        background: 'rgba(13, 148, 136, 0.1)',
        color: 'var(--pm-primary, #0d9488)',
      },
    },
  },
  content: {
    paddingTop: '20px',
    paddingLeft: '32px',
    paddingRight: '32px',
  },
  footerInner: {
    background: '#f8fafc',
    borderTop: '1px solid #e2e8f0',
    padding: '12px 24px',
  },
  contentInner: {
    // Ensure the content area fills available space
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
  },
};

export const StyledPanel: React.FC<IPanelProps> = (props) => {
  // Merge caller's styles with our standard styles
  const mergedStyles: Partial<IPanelStyles> = { ...PANEL_STYLES };

  if (props.styles) {
    const callerStyles = typeof props.styles === 'function'
      ? props.styles({} as any)
      : props.styles;

    if (callerStyles) {
      Object.keys(callerStyles).forEach(key => {
        const k = key as keyof IPanelStyles;
        if (mergedStyles[k] && callerStyles[k]) {
          (mergedStyles as any)[k] = { ...(mergedStyles[k] as any), ...(callerStyles[k] as any) };
        } else if (callerStyles[k]) {
          (mergedStyles as any)[k] = callerStyles[k];
        }
      });
    }
  }

  return <Panel {...props} styles={mergedStyles} />;
};

export default StyledPanel;
