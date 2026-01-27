/**
 * Fluent UI v9 Portal Styles Injector
 *
 * Uses MutationObserver to watch for portal elements and apply inline styles
 * directly, which has higher specificity than Griffel's atomic CSS classes.
 *
 * This fixes floating backgrounds on Fluent UI v9 portal components
 * (Dropdown, Dialog, Tooltip, Menu, Popover, Drawer).
 */

const PORTAL_OBSERVER_ID = 'jml-portal-observer';
let observer: MutationObserver | null = null;
let isInitialized = false;

// Fluent UI font stack
const FLUENT_FONT_FAMILY = "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";

// Style configurations for different portal elements - Modern Fluent Design
const portalStyles: Record<string, Record<string, string>> = {
  // Listbox (Dropdown options container)
  '[role="listbox"]': {
    backgroundColor: '#ffffff',
    border: '1px solid #e1dfdd',
    boxShadow: '0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08)',
    borderRadius: '8px',
    padding: '4px',
    overflow: 'hidden',
    fontFamily: FLUENT_FONT_FAMILY
  },
  // Option items
  '[role="option"]': {
    backgroundColor: '#ffffff',
    color: '#323130',
    padding: '8px 12px',
    cursor: 'pointer',
    minHeight: '36px',
    display: 'flex',
    alignItems: 'center',
    borderRadius: '4px',
    margin: '2px 4px',
    fontFamily: FLUENT_FONT_FAMILY,
    fontSize: '14px'
  },
  // Dialog surface
  '[role="dialog"]': {
    backgroundColor: '#ffffff',
    borderRadius: '12px',
    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.14), 0 2px 8px rgba(0, 0, 0, 0.12)',
    color: '#323130',
    border: '1px solid rgba(0, 0, 0, 0.05)',
    fontFamily: FLUENT_FONT_FAMILY
  },
  // Tooltip
  '[role="tooltip"]': {
    backgroundColor: '#323130',
    color: '#ffffff',
    borderRadius: '6px',
    boxShadow: '0 4px 12px rgba(0, 0, 0, 0.15)',
    padding: '8px 12px',
    fontSize: '12px',
    fontFamily: FLUENT_FONT_FAMILY
  },
  // Menu
  '[role="menu"]': {
    backgroundColor: '#ffffff',
    border: '1px solid #e1dfdd',
    borderRadius: '8px',
    boxShadow: '0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08)',
    padding: '4px',
    fontFamily: FLUENT_FONT_FAMILY
  },
  // Menu items
  '[role="menuitem"]': {
    backgroundColor: '#ffffff',
    color: '#323130',
    padding: '8px 12px',
    borderRadius: '4px',
    margin: '2px 4px',
    fontFamily: FLUENT_FONT_FAMILY,
    fontSize: '14px'
  },
  // Alertdialog
  '[role="alertdialog"]': {
    backgroundColor: '#ffffff',
    borderRadius: '12px',
    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.14), 0 2px 8px rgba(0, 0, 0, 0.12)',
    color: '#323130',
    border: '1px solid rgba(0, 0, 0, 0.05)',
    fontFamily: FLUENT_FONT_FAMILY
  },
  // Combobox
  '[role="combobox"]': {
    backgroundColor: '#ffffff',
    border: '1px solid #d1d1d1',
    borderRadius: '6px',
    color: '#323130',
    fontFamily: FLUENT_FONT_FAMILY
  }
};

/**
 * Apply inline styles to an element
 */
function applyStyles(element: HTMLElement, styles: Record<string, string>): void {
  Object.entries(styles).forEach(([property, value]) => {
    element.style.setProperty(property.replace(/([A-Z])/g, '-$1').toLowerCase(), value, 'important');
  });
}

/**
 * Style a single portal element based on its role
 */
function stylePortalElement(element: HTMLElement): void {
  const role = element.getAttribute('role');

  if (role && portalStyles[`[role="${role}"]`]) {
    applyStyles(element, portalStyles[`[role="${role}"]`]);
    element.setAttribute('data-jml-styled', 'true');
  }
}

/**
 * Find and style all portal elements in a subtree
 */
function styleAllPortalElements(root: Element | Document = document): void {
  // Style elements by role
  const roles = ['listbox', 'option', 'dialog', 'alertdialog', 'tooltip', 'menu', 'menuitem'];

  roles.forEach(role => {
    const elements = root.querySelectorAll(`[role="${role}"]:not([data-jml-styled])`);
    elements.forEach(element => {
      if (element instanceof HTMLElement) {
        stylePortalElement(element);
      }
    });
  });

  // Also style by Fluent UI class patterns
  const fluentSelectors = [
    '.fui-Listbox',
    '.fui-Option',
    '.fui-DialogSurface',
    '.fui-DialogBody',
    '.fui-DialogContent',
    '.fui-DialogActions',
    '.fui-Tooltip',
    '.fui-PopoverSurface',
    '.fui-MenuList',
    '.fui-MenuItem',
    '.fui-MenuPopover',
    '.fui-DrawerSurface',
    '.fui-DrawerBody'
  ];

  fluentSelectors.forEach(selector => {
    const elements = root.querySelectorAll(`${selector}:not([data-jml-styled])`);
    elements.forEach(element => {
      if (element instanceof HTMLElement) {
        const className = selector.replace('.fui-', '').toLowerCase();

        // Apply appropriate styles based on class
        if (className.includes('listbox')) {
          applyStyles(element, portalStyles['[role="listbox"]']);
        } else if (className.includes('option')) {
          applyStyles(element, portalStyles['[role="option"]']);
        } else if (className.includes('dialog') || className.includes('drawer')) {
          applyStyles(element, {
            backgroundColor: '#ffffff',
            color: '#323130',
            fontFamily: FLUENT_FONT_FAMILY
          });
        } else if (className.includes('tooltip')) {
          applyStyles(element, portalStyles['[role="tooltip"]']);
        } else if (className.includes('menu') || className.includes('popover')) {
          applyStyles(element, {
            backgroundColor: '#ffffff',
            border: '1px solid #d1d1d1',
            borderRadius: '4px',
            boxShadow: '0 6.4px 14.4px 0 rgba(0, 0, 0, 0.132), 0 1.2px 3.6px 0 rgba(0, 0, 0, 0.108)',
            color: '#323130',
            fontFamily: FLUENT_FONT_FAMILY
          });
        }

        element.setAttribute('data-jml-styled', 'true');
      }
    });
  });
}

/**
 * Callback for MutationObserver
 */
function handleMutations(mutations: MutationRecord[]): void {
  mutations.forEach(mutation => {
    // Handle added nodes
    mutation.addedNodes.forEach(node => {
      if (node instanceof HTMLElement) {
        // Check if the node itself needs styling
        stylePortalElement(node);
        // Check all descendants
        styleAllPortalElements(node);
      }
    });

    // Handle attribute changes (for dynamically set roles)
    if (mutation.type === 'attributes' && mutation.attributeName === 'role') {
      if (mutation.target instanceof HTMLElement) {
        stylePortalElement(mutation.target);
      }
    }
  });
}

/**
 * Initialize the portal styles observer.
 * Safe to call multiple times - will only initialize once.
 */
export function injectPortalStyles(): void {
  if (isInitialized) {
    return;
  }

  // Check if already initialized by another webpart instance
  if (document.body.hasAttribute(PORTAL_OBSERVER_ID)) {
    isInitialized = true;
    return;
  }

  // Mark as initialized
  document.body.setAttribute(PORTAL_OBSERVER_ID, 'true');

  // Style any existing portal elements
  styleAllPortalElements();

  // Create MutationObserver to watch for new portal elements
  observer = new MutationObserver(handleMutations);

  // Observe the entire document for changes
  observer.observe(document.body, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['role', 'class']
  });

  // Also inject CSS as a fallback for elements we might miss
  injectFallbackCSS();

  // Inject SharePoint overrides to hide social bar, comments, etc.
  injectSharePointOverrides();

  isInitialized = true;
  console.log('[JML] Portal styles observer initialized');
}

/**
 * Inject fallback CSS styles
 */
function injectFallbackCSS(): void {
  const styleId = 'jml-portal-fallback-css';
  if (document.getElementById(styleId)) {
    return;
  }

  const css = `
    /* ===========================================
       JML Portal Styles - Modern Fluent Design
       Fixes floating/transparent backgrounds
       =========================================== */

    /* ----- Dialog Overlay/Backdrop ----- */
    .fui-DialogSurface__backdrop,
    [class*="DialogSurface__backdrop"],
    div[style*="position: fixed"][style*="inset: 0"],
    div[style*="position: fixed"][style*="top: 0"][style*="left: 0"][style*="right: 0"][style*="bottom: 0"] {
      background-color: rgba(0, 0, 0, 0.4) !important;
    }

    /* ----- Dialog Surface ----- */
    [role="dialog"],
    [role="alertdialog"],
    .fui-DialogSurface {
      background-color: #ffffff !important;
      color: #323130 !important;
      border-radius: 12px !important;
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.14), 0 2px 8px rgba(0, 0, 0, 0.12) !important;
      border: 1px solid rgba(0, 0, 0, 0.05) !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-DialogBody,
    .fui-DialogContent,
    .fui-DialogActions,
    .fui-DialogTitle {
      background-color: #ffffff !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- Dropdown/Listbox ----- */
    [role="listbox"],
    .fui-Listbox {
      background-color: #ffffff !important;
      border: 1px solid #e1dfdd !important;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08) !important;
      border-radius: 8px !important;
      padding: 4px !important;
      overflow: hidden !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    [role="option"],
    .fui-Option {
      background-color: #ffffff !important;
      color: #323130 !important;
      border-radius: 4px !important;
      margin: 2px 4px !important;
      padding: 8px 12px !important;
      min-height: 36px !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
      font-size: 14px !important;
    }

    [role="option"]:hover,
    .fui-Option:hover {
      background-color: #f5f5f5 !important;
    }

    [role="option"][aria-selected="true"],
    .fui-Option[aria-selected="true"] {
      background-color: #e5f1fb !important;
      color: #0078d4 !important;
    }

    [role="option"]:focus,
    .fui-Option:focus {
      background-color: #f5f5f5 !important;
      outline: 2px solid #0078d4 !important;
      outline-offset: -2px !important;
    }

    /* ----- Combobox/Dropdown Trigger ----- */
    .fui-Combobox,
    .fui-Dropdown {
      background-color: #ffffff !important;
      border: 1px solid #d1d1d1 !important;
      border-radius: 6px !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Combobox:hover,
    .fui-Dropdown:hover {
      border-color: #b3b0ad !important;
    }

    .fui-Combobox:focus-within,
    .fui-Dropdown:focus-within {
      border-color: #0078d4 !important;
      box-shadow: 0 0 0 1px #0078d4 !important;
    }

    .fui-Combobox__input,
    .fui-Dropdown__button {
      background-color: transparent !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- Tooltip ----- */
    [role="tooltip"],
    .fui-Tooltip {
      background-color: #323130 !important;
      color: #ffffff !important;
      border-radius: 6px !important;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
      padding: 8px 12px !important;
      font-size: 12px !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- Menu ----- */
    [role="menu"],
    .fui-MenuList,
    .fui-MenuPopover {
      background-color: #ffffff !important;
      border: 1px solid #e1dfdd !important;
      border-radius: 8px !important;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08) !important;
      padding: 4px !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    [role="menuitem"],
    .fui-MenuItem {
      background-color: #ffffff !important;
      color: #323130 !important;
      border-radius: 4px !important;
      margin: 2px 4px !important;
      padding: 8px 12px !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
      font-size: 14px !important;
    }

    [role="menuitem"]:hover,
    .fui-MenuItem:hover {
      background-color: #f5f5f5 !important;
    }

    [role="menuitem"]:focus,
    .fui-MenuItem:focus {
      background-color: #f5f5f5 !important;
      outline: none !important;
    }

    .fui-MenuDivider {
      background-color: #e1dfdd !important;
      margin: 4px 8px !important;
    }

    /* ----- Popover ----- */
    .fui-PopoverSurface {
      background-color: #ffffff !important;
      border: 1px solid #e1dfdd !important;
      border-radius: 8px !important;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08) !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- Drawer ----- */
    /* JML STANDARD: Square corners on all panels/drawers */
    .fui-DrawerSurface,
    .fui-DrawerBody,
    .fui-DrawerHeader,
    .fui-DrawerHeaderTitle,
    .fui-DrawerFooter,
    [class*="fui-DrawerSurface"],
    [class*="fui-Drawer"] > div {
      background-color: #ffffff !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
      border-radius: 0 !important;
      border-top-left-radius: 0 !important;
      border-top-right-radius: 0 !important;
      border-bottom-left-radius: 0 !important;
      border-bottom-right-radius: 0 !important;
    }

    .fui-DrawerSurface,
    [class*="fui-DrawerSurface"] {
      box-shadow: -4px 0 24px rgba(0, 0, 0, 0.14) !important;
      border-radius: 0 !important;
    }

    /* ----- Input/TextField ----- */
    .fui-Input,
    .fui-Textarea {
      background-color: #ffffff !important;
      border: 1px solid #d1d1d1 !important;
      border-radius: 6px !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Input:hover,
    .fui-Textarea:hover {
      border-color: #b3b0ad !important;
    }

    .fui-Input:focus-within,
    .fui-Textarea:focus-within {
      border-color: #0078d4 !important;
      box-shadow: 0 0 0 1px #0078d4 !important;
    }

    /* ----- Card ----- */
    .fui-Card {
      background-color: #ffffff !important;
      border: 1px solid #e1dfdd !important;
      border-radius: 8px !important;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.04) !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Card:hover {
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
    }

    /* ----- Portal Z-Index Fix ----- */
    .fui-Portal,
    [data-portal-node] {
      z-index: 1000000 !important;
    }

    /* ----- Fix for position fixed containers ----- */
    .fui-FluentProvider > div[style*="position: fixed"],
    body > div[style*="position: fixed"]:not([class*="backdrop"]) {
      /* Allow pointer events but keep transparent */
    }

    /* ----- Spinner/Loading ----- */
    .fui-Spinner {
      color: #0078d4 !important;
    }

    /* ----- MessageBar ----- */
    .fui-MessageBar {
      background-color: #ffffff !important;
      border-radius: 6px !important;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.04) !important;
    }

    /* ----- Tab List ----- */
    .fui-TabList {
      background-color: transparent !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Tab {
      color: #605e5c !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Tab[aria-selected="true"] {
      color: #0078d4 !important;
    }

    /* ----- Badge ----- */
    .fui-Badge {
      font-weight: 600 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- Persona ----- */
    .fui-Avatar,
    .fui-Persona {
      background-color: #f3f2f1 !important;
    }

    /* ----- Table ----- */
    .fui-Table {
      background-color: #ffffff !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-TableRow:hover {
      background-color: #f5f5f5 !important;
    }

    .fui-TableHeaderCell {
      background-color: #fafafa !important;
      border-bottom: 1px solid #e1dfdd !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-TableCell {
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- Accordion ----- */
    .fui-AccordionHeader {
      background-color: #ffffff !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-AccordionPanel {
      background-color: #ffffff !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- InfoLabel Popover ----- */
    .fui-InfoLabel__popover {
      background-color: #ffffff !important;
      border: 1px solid #e1dfdd !important;
      border-radius: 8px !important;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08) !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    /* ----- DatePicker Popover ----- */
    .fui-DatePicker__popover,
    .fui-Calendar {
      background-color: #ffffff !important;
      border: 1px solid #e1dfdd !important;
      border-radius: 8px !important;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.14), 0 2px 6px rgba(0, 0, 0, 0.08) !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Calendar__day {
      background-color: #ffffff !important;
      color: #323130 !important;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif !important;
    }

    .fui-Calendar__day:hover {
      background-color: #f5f5f5 !important;
    }

    .fui-Calendar__day[aria-selected="true"] {
      background-color: #0078d4 !important;
      color: #ffffff !important;
    }
  `;

  const style = document.createElement('style');
  style.id = styleId;
  style.textContent = css;
  document.head.appendChild(style);
}

/**
 * Inject CSS to hide SharePoint OOB social bar and page elements
 * This ensures consistent JML branding across all pages
 */
function injectSharePointOverrides(): void {
  const styleId = 'jml-sharepoint-overrides-css';
  if (document.getElementById(styleId)) {
    return;
  }

  const css = `
    /* ===========================================
       JML SharePoint Overrides
       Hides OOB social bar, comments, reactions
       Based on JML Official Style Guide
       =========================================== */

    /* ----- Hide SharePoint Social Bar & Comments ----- */
    /* Like, Views, Save for later buttons */
    [data-automation-id="pageCommands"],
    [data-automation-id="socialBar"],
    [data-automation-id="PageLikes"],
    [data-automation-id="pageComments"],
    [data-automation-id="likeButton"],
    [data-automation-id="viewsCount"],
    [data-automation-id="saveForLater"],
    [data-automation-id="pageLikeButton"],
    [data-automation-id="pageViewCount"],
    [data-automation-id="pageSaveForLater"],
    [data-automation-id="pageReactions"],
    [data-automation-id="PageSocialBar"],
    [data-automation-id="pageFooter"],
    [data-automation-id="SiteFooter"],
    .pageCommands,
    .ms-HubLinks,
    .ms-CommentsWrapper {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      overflow: hidden !important;
    }

    /* Class-based selectors for social bar */
    div[class*="socialBar"],
    div[class*="SocialBar"],
    div[class*="pageCommandsWrapper"],
    div[class*="PageCommands"],
    div[class*="pageActions"],
    div[class*="PageActions"],
    div[class*="pageSocialBar"],
    div[class*="pageReactions"],
    div[class*="PageReactions"],
    div[class*="pageLikes"],
    div[class*="PageLikes"],
    div[class*="CommentsWrapper"],
    div[class*="likeButton"],
    div[class*="LikeButton"],
    div[class*="viewCount"],
    div[class*="ViewCount"],
    div[class*="saveForLater"],
    div[class*="SaveForLater"],
    div[class*="pageInteractions"],
    div[class*="PageInteractions"],
    div[class*="pageMetrics"],
    div[class*="PageMetrics"],
    div[class*="pageFooterActions"],
    div[class*="PageFooterActions"],
    div[class*="pageCommandBar"],
    div[class*="PageCommandBar"],
    div[class*="actionsBar"],
    div[class*="ActionsBar"],
    div[class*="pageInteractionBar"],
    div[class*="PageInteractionBar"],
    div[class*="engagement"],
    div[class*="Engagement"],
    div[class*="commentsContainer"],
    div[class*="CommentsContainer"],
    div[data-sp-feature-tag*="PageMetadata"],
    div[class*="canvasPageMetadata"],
    div[class*="CanvasPageMetadata"],
    [class*="pageActions_"],
    [class*="PageActions_"],
    [class*="pageCommandsWrapper_"],
    [class*="PageCommandsWrapper_"],
    [class*="pageMetadata_"],
    [class*="PageMetadata_"],
    [class*="pageMetadataRow"],
    [class*="PageMetadataRow"],
    section[class*="pageReactionsContainer"],
    section[class*="PageReactionsContainer"],
    div[class*="reactionBar"],
    div[class*="ReactionBar"] {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      overflow: hidden !important;
    }

    /* Hide page rating/feedback */
    div[class*="pageRating"],
    div[class*="pageFeedback"],
    [data-automation-id="pageRating"],
    [class*="pageBottomCommands"],
    [class*="PageBottomCommands"],
    [class*="feedbackContainer"],
    [class*="FeedbackContainer"],
    [class*="pageRatingContainer"],
    [class*="PageRatingContainer"] {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      overflow: hidden !important;
    }

    /* ----- Hide Social Bar Separators/Dividers ----- */
    [class*="pageSeparator"],
    [class*="PageSeparator"],
    [class*="socialBarSeparator"],
    [class*="SocialBarSeparator"],
    [class*="pageCommandsSeparator"],
    [class*="PageCommandsSeparator"],
    [class*="pageDivider"],
    [class*="PageDivider"],
    [class*="contentSeparator"],
    [class*="ContentSeparator"],
    [data-automation-id="pageSeparator"],
    [data-automation-id="PageSeparator"],
    [data-automation-id="contentDivider"],
    hr[class*="separator"],
    hr[class*="Separator"],
    hr[class*="divider"],
    hr[class*="Divider"] {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      border: none !important;
    }

    /* Remove borders on social bar parent containers */
    [class*="pageMetadata"],
    [class*="PageMetadata"],
    [class*="pageCommands"],
    [class*="PageCommands"],
    [class*="socialBar"],
    [class*="SocialBar"],
    [class*="pageActionsContainer"],
    [class*="PageActionsContainer"] {
      border: none !important;
      border-top: none !important;
      border-bottom: none !important;
      box-shadow: none !important;
    }

    /* ----- Aggressive Border Removal on Layout Containers ----- */
    /* Remove ALL borders from SharePoint canvas/layout containers */
    [class*="CanvasZone"],
    [class*="canvasZone"],
    [class*="CanvasSection"],
    [class*="canvasSection"],
    [class*="ControlZone"],
    [class*="controlZone"],
    [class*="WebPartZone"],
    [class*="webPartZone"],
    [class*="pageContent"],
    [class*="PageContent"],
    [class*="mainContent"],
    [class*="MainContent"],
    [class*="contentBox"],
    [class*="ContentBox"],
    [class*="pageRegion"],
    [class*="PageRegion"],
    [class*="articleRegion"],
    [class*="ArticleRegion"],
    [class*="contentRegion"],
    [class*="ContentRegion"],
    [class*="bottomRegion"],
    [class*="BottomRegion"],
    [class*="pageBody"],
    [class*="PageBody"],
    [class*="SPPageChrome"],
    [class*="spPageChrome"],
    [data-automation-id="contentScrollRegion"],
    [data-automation-id="CanvasZone"],
    [data-automation-id="CanvasSection"],
    .mainContent,
    .pageContent,
    #contentBox,
    #contentRow,
    article,
    main {
      border: none !important;
      border-top: none !important;
      border-bottom: none !important;
      box-shadow: none !important;
    }

    /* Hide any remaining HR elements */
    hr:not([data-sp-web-part] *):not(.ms-Panel *):not(.ms-Layer *) {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      border: none !important;
    }

    /* ----- AGGRESSIVE: Hide thin separator lines by structure ----- */
    /* Target thin divs that are likely separators (often dynamically classed) */
    #spPageCanvasContent > div:not([class*="Canvas"]):not([class*="Zone"]):not([data-sp-web-part]):not(.ms-Panel *),
    #spPageCanvasContent + div:not([class*="Canvas"]):not([data-sp-web-part]):not(.ms-Panel *) {
      border: none !important;
      border-top: none !important;
      border-bottom: none !important;
    }

    /* Hide any empty divs at the page bottom that might be separators */
    body > div:empty:not(.ms-Layer):not(.ms-Panel):not(.ms-LayerHost),
    #spPageCanvasContent ~ div:empty:not(.ms-Layer):not(.ms-Panel) {
      display: none !important;
    }

    /* Remove borders from all content wrappers */
    [class*="pageContent"]:not(.ms-Panel *),
    [class*="PageContent"]:not(.ms-Panel *),
    [class*="articleContent"]:not(.ms-Panel *),
    [class*="ArticleContent"]:not(.ms-Panel *) {
      border: none !important;
      border-top: none !important;
      border-bottom: none !important;
    }

    /* ----- Hide SharePoint Page Footer COMPLETELY ----- */
    /* JML has its own custom footer component inside webparts */
    /* Hide ALL SharePoint footers that are NOT inside webparts */
    footer:not([data-sp-web-part] *):not(.ms-Panel *):not(.ms-Layer *):not([class*="jml"]):not([class*="Jml"]):not([class*="JML"]):not([data-jml-footer]),
    footer[data-automation-id="pageFooter"],
    footer[data-automation-id="SiteFooter"],
    footer.ms-compositeFooter,
    .ms-siteFooter,
    #spSiteFooter,
    div[class*="siteFooter"]:not([data-sp-web-part] *):not([class*="jml"]):not(.ms-Panel *),
    div[class*="SiteFooter"]:not([data-sp-web-part] *):not([class*="jml"]):not(.ms-Panel *),
    footer[class*="compositeFooter"]:not([data-sp-web-part] *),
    footer[class*="CompositeFooter"]:not([data-sp-web-part] *),
    div[class*="siteLevelFooter"]:not([data-sp-web-part] *),
    div[class*="SiteLevelFooter"]:not([data-sp-web-part] *),
    footer[class*="footer_"]:not([data-sp-web-part] *):not(.ms-Panel *),
    footer[class*="Footer_"]:not([data-sp-web-part] *):not(.ms-Panel *),
    [class*="pageFooter"]:not([data-sp-web-part] *):not(.ms-Panel *),
    [class*="PageFooter"]:not([data-sp-web-part] *):not(.ms-Panel *),
    [class*="page-footer"]:not([data-sp-web-part] *):not(.ms-Panel *) {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      max-height: 0 !important;
      overflow: hidden !important;
      padding: 0 !important;
      margin: 0 !important;
      border: none !important;
    }

    /* ----- AGGRESSIVE Social Bar Hiding ----- */
    /* Target ANY element that contains social bar content */
    /* Social bar container patterns */
    [class*="likesAndComments"]:not(.ms-Panel *),
    [class*="LikesAndComments"]:not(.ms-Panel *),
    [class*="pageSocialSection"]:not(.ms-Panel *),
    [class*="PageSocialSection"]:not(.ms-Panel *),
    [class*="socialContainer"]:not(.ms-Panel *),
    [class*="SocialContainer"]:not(.ms-Panel *),
    /* Buttons by aria-label */
    button[aria-label="Like"]:not(.ms-Panel *),
    button[aria-label="Save for later"]:not(.ms-Panel *),
    button[aria-label*="View"]:not(.ms-Panel *):not([data-sp-web-part] *),
    /* Parent containers of Like/Save buttons */
    div:has(> button[aria-label="Like"]):not([data-sp-web-part] *):not(.ms-Panel *),
    section:has(> button[aria-label="Like"]):not([data-sp-web-part] *):not(.ms-Panel *),
    /* Views counter (text that shows view count) */
    span[class*="viewCount"]:not(.ms-Panel *),
    span[class*="ViewCount"]:not(.ms-Panel *),
    div[class*="viewCount"]:not(.ms-Panel *),
    div[class*="ViewCount"]:not(.ms-Panel *) {
      display: none !important;
      visibility: hidden !important;
      height: 0 !important;
      overflow: hidden !important;
    }

    /* Ensure JML footers inside webparts are visible */
    [data-sp-web-part] footer,
    [data-sp-web-part] [class*="jml"][class*="footer"],
    [data-sp-web-part] [class*="Jml"][class*="Footer"],
    [data-sp-web-part] [class*="JML"][class*="Footer"] {
      display: block !important;
      visibility: visible !important;
      height: auto !important;
      overflow: visible !important;
    }

    /* ----- Full-Width Support ----- */
    /* Allow content to extend full width when needed */
    .CanvasZone,
    [class*="CanvasZone"],
    .CanvasSection,
    [class*="CanvasSection"],
    .ControlZone,
    [class*="ControlZone"],
    .CanvasZoneContainer,
    [class*="CanvasZoneContainer"],
    [data-sp-web-part],
    .webPartContainer,
    [class*="webPartContainer"],
    [class*="CanvasComponent"],
    .CanvasComponent {
      overflow: visible !important;
      max-width: none !important;
    }
  `;

  const style = document.createElement('style');
  style.id = styleId;
  style.textContent = css;
  document.head.appendChild(style);

  // Also use JavaScript to directly hide elements (more reliable than CSS alone)
  hideSharePointElements();

  // Set up observer to catch dynamically loaded elements
  setupSharePointElementObserver();

  console.log('[JML] SharePoint overrides CSS injected');
}

/**
 * Selectors for SharePoint elements to hide
 */
const SP_HIDE_SELECTORS = [
  // Social bar elements
  '[data-automation-id="pageCommands"]',
  '[data-automation-id="socialBar"]',
  '[data-automation-id="PageLikes"]',
  '[data-automation-id="pageComments"]',
  '[data-automation-id="PageSocialBar"]',
  '[data-automation-id="pageFooter"]',
  '[data-automation-id="SiteFooter"]',
  // Footer elements
  'footer[data-automation-id="pageFooter"]',
  'footer[data-automation-id="SiteFooter"]',
  'footer.ms-compositeFooter',
  '.ms-siteFooter',
  '#spSiteFooter',
  // Class-based selectors
  '[class*="socialBar_"]',
  '[class*="SocialBar_"]',
  '[class*="pageCommands_"]',
  '[class*="PageCommands_"]',
  '[class*="siteFooter_"]',
  '[class*="SiteFooter_"]',
  '[class*="compositeFooter"]'
];

/**
 * Check if an element is inside a Fluent UI Panel, Layer, or Overlay
 * These should NEVER be hidden as they're used for fly-in panels, dialogs, etc.
 */
function isInsidePanel(el: Element): boolean {
  // Check if element is inside any Panel/Layer/Overlay structure
  if (el.closest('.ms-Panel') ||
      el.closest('.ms-Layer') ||
      el.closest('.ms-Overlay') ||
      el.closest('[class*="ms-Panel"]') ||
      el.closest('[class*="ms-Layer"]') ||
      el.closest('[role="dialog"]') ||
      el.closest('[role="alertdialog"]') ||
      el.closest('.fui-DialogSurface') ||
      el.closest('.fui-DrawerSurface')) {
    return true;
  }

  // Check if the element itself is a Panel/Layer component
  if (el.classList) {
    const classList = el.classList.toString();
    if (classList.includes('ms-Panel') ||
        classList.includes('ms-Layer') ||
        classList.includes('ms-Overlay') ||
        classList.includes('fui-Dialog') ||
        classList.includes('fui-Drawer')) {
      return true;
    }
  }

  return false;
}

/**
 * Hide SharePoint elements using direct DOM manipulation
 * IMPORTANT: Skip elements inside webparts and Panels to preserve JML UI
 */
function hideSharePointElements(): void {
  SP_HIDE_SELECTORS.forEach(selector => {
    try {
      const elements = document.querySelectorAll(selector);
      elements.forEach(el => {
        // Skip if inside a webpart (preserves JML footers and other UI)
        if (el.closest('[data-sp-web-part]')) {
          return;
        }

        // Skip if inside a Panel, Layer, or Overlay (CRITICAL - fixes Panel display issue)
        if (isInsidePanel(el)) {
          return;
        }

        // Skip if has JML class
        if (el instanceof HTMLElement &&
            (el.classList.toString().toLowerCase().includes('jml') ||
             el.id?.toLowerCase().includes('jml'))) {
          return;
        }

        if (el instanceof HTMLElement && !el.hasAttribute('data-jml-hidden')) {
          el.style.setProperty('display', 'none', 'important');
          el.style.setProperty('visibility', 'hidden', 'important');
          el.style.setProperty('height', '0', 'important');
          el.style.setProperty('overflow', 'hidden', 'important');
          el.setAttribute('data-jml-hidden', 'true');
        }
      });
    } catch (e) {
      // Ignore selector errors
    }
  });

  // Also hide elements containing "Like" or "Save for later" buttons
  hideByButtonContent();
}

/**
 * Hide containers that have Like/Save buttons by checking button content
 * CONSERVATIVE approach - only hide small containers, max 4 levels up
 */
function hideByButtonContent(): void {
  // Find buttons with EXACT aria-labels (not partial matches to avoid false positives)
  const socialButtonSelectors = [
    'button[aria-label="Like"]',
    'button[aria-label="Save for later"]'
  ];

  socialButtonSelectors.forEach(selector => {
    try {
      const buttons = document.querySelectorAll(selector);
      buttons.forEach(btn => {
        // Skip if inside a webpart
        if (btn.closest('[data-sp-web-part]')) {
          return;
        }

        // Skip if inside a Panel (CRITICAL - fixes Panel display issue)
        if (isInsidePanel(btn)) {
          return;
        }

        // Only go up 4 levels max to find container
        let parent = btn.parentElement;
        for (let i = 0; i < 4 && parent; i++) {
          // Skip if parent is inside a Panel
          if (isInsidePanel(parent)) {
            break;
          }

          // Check if this container has BOTH Like AND Save buttons (the social bar)
          const hasLike = parent.querySelector('button[aria-label="Like"]');
          const hasSave = parent.querySelector('button[aria-label="Save for later"]');

          if (hasLike && hasSave) {
            // Safety check: only hide if element height is under 100px (social bar is small)
            const rect = parent.getBoundingClientRect();
            if (rect.height < 100 && !parent.hasAttribute('data-jml-hidden')) {
              (parent as HTMLElement).style.setProperty('display', 'none', 'important');
              parent.setAttribute('data-jml-hidden', 'true');
              console.log('[JML] Hidden social bar container via button detection');
            }
            break;
          }
          parent = parent.parentElement;
        }
      });
    } catch (e) {
      // Ignore selector errors
    }
  });

  // Look for specific SharePoint social bar elements
  hideSocialBarByTextContent();

  // Hide all HR elements and separator lines outside of webparts
  hideSeparatorLines();

  // Hide SharePoint page footers (not JML footers or webpart footers)
  const footers = document.querySelectorAll('footer');
  footers.forEach(footer => {
    // Skip if already hidden
    if (footer.hasAttribute('data-jml-hidden')) {
      return;
    }

    // Skip if inside a webpart (JML footers will be inside webparts)
    if (footer.closest('[data-sp-web-part]')) {
      return;
    }

    // Skip if inside a Panel (CRITICAL - fixes Panel display issue)
    if (isInsidePanel(footer)) {
      return;
    }

    // Skip if this is a JML footer (has jml class or data attribute)
    if (footer.classList.toString().toLowerCase().includes('jml') ||
        footer.hasAttribute('data-jml-footer') ||
        footer.id?.toLowerCase().includes('jml')) {
      return;
    }

    // Hide SharePoint footer
    (footer as HTMLElement).style.setProperty('display', 'none', 'important');
    footer.setAttribute('data-jml-hidden', 'true');
    console.log('[JML] Hidden SharePoint footer');
  });
}

/**
 * Hide social bar by looking for specific SharePoint social bar patterns
 * CONSERVATIVE approach - only hide small, specific elements
 */
function hideSocialBarByTextContent(): void {
  // Only look for elements with specific SharePoint social bar class patterns
  // These are much more specific and won't accidentally hide page content
  const socialBarSelectors = [
    // SharePoint specific social bar classes
    '[class*="root_"][class*="pageMetadata"]',
    '[class*="pageMetadataContainer"]',
    '[class*="pageSocialBar"]',
    '[class*="PageSocialBar"]',
    '[class*="pageActionsContainer"]',
    '[class*="PageActionsContainer"]',
    // Specific data attributes
    '[data-automation-id*="social"]',
    '[data-automation-id*="Social"]',
    '[data-automation-id*="like"]',
    '[data-automation-id*="Like"]',
    '[data-automation-id*="pageCommand"]',
    '[data-automation-id*="PageCommand"]'
  ];

  socialBarSelectors.forEach(selector => {
    try {
      const elements = document.querySelectorAll(selector);
      elements.forEach(el => {
        // Skip if inside a Panel (CRITICAL - fixes Panel display issue)
        if (isInsidePanel(el)) {
          return;
        }

        if (!el.hasAttribute('data-jml-hidden') && !el.closest('[data-sp-web-part]')) {
          // Extra safety: only hide if element is reasonably small (less than 200px tall)
          const rect = el.getBoundingClientRect();
          if (rect.height < 200) {
            (el as HTMLElement).style.setProperty('display', 'none', 'important');
            el.setAttribute('data-jml-hidden', 'true');
            console.log('[JML] Hidden social bar element:', selector);
          }
        }
      });
    } catch (e) {
      // Ignore selector errors
    }
  });
}

/**
 * Hide separator lines (HR elements and thin dividers) outside of webparts
 * Also removes any visible borders that could look like separators
 */
function hideSeparatorLines(): void {
  // Hide all HR elements not inside webparts or panels
  const hrElements = document.querySelectorAll('hr');
  hrElements.forEach(hr => {
    // Skip if inside a Panel (CRITICAL - fixes Panel display issue)
    if (isInsidePanel(hr)) {
      return;
    }

    if (!hr.hasAttribute('data-jml-hidden') && !hr.closest('[data-sp-web-part]')) {
      (hr as HTMLElement).style.setProperty('display', 'none', 'important');
      (hr as HTMLElement).style.setProperty('visibility', 'hidden', 'important');
      (hr as HTMLElement).style.setProperty('height', '0', 'important');
      (hr as HTMLElement).style.setProperty('border', 'none', 'important');
      hr.setAttribute('data-jml-hidden', 'true');
      console.log('[JML] Hidden HR element');
    }
  });

  // Find thin elements that look like separator lines (height < 5px, width > 100px)
  // Check divs and spans that might be used as visual separators
  // IMPORTANT: Exclude elements inside Panels to avoid breaking Panel overlays
  const potentialSeparators = document.querySelectorAll('div:not([data-sp-web-part] *), span:not([data-sp-web-part] *)');
  potentialSeparators.forEach(el => {
    // Skip if inside a Panel (CRITICAL - fixes Panel display issue)
    if (isInsidePanel(el)) {
      return;
    }

    if (el.hasAttribute('data-jml-hidden') || el.closest('[data-sp-web-part]')) {
      return;
    }

    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);

    // Check if element looks like a separator line:
    // - Very thin (height 1-5px)
    // - Relatively wide (> 200px)
    // - Has a background color or border that makes it visible
    const isThinLine = rect.height > 0 && rect.height <= 5 && rect.width > 200;
    const hasVisibleBackground = style.backgroundColor !== 'rgba(0, 0, 0, 0)' && style.backgroundColor !== 'transparent';
    const hasVisibleBorder = style.borderTopWidth !== '0px' || style.borderBottomWidth !== '0px';

    if (isThinLine && (hasVisibleBackground || hasVisibleBorder)) {
      (el as HTMLElement).style.setProperty('display', 'none', 'important');
      el.setAttribute('data-jml-hidden', 'true');
      console.log('[JML] Hidden thin separator element');
    }
  });

  // Remove top/bottom borders from canvas containers (where separator might come from)
  // IMPORTANT: Exclude elements inside Panels
  const canvasContainers = document.querySelectorAll('[class*="Canvas"], [class*="canvas"], [class*="Region"], [class*="region"]');
  canvasContainers.forEach(el => {
    // Skip if inside a Panel
    if (isInsidePanel(el)) {
      return;
    }

    if (!el.closest('[data-sp-web-part]')) {
      (el as HTMLElement).style.setProperty('border-top', 'none', 'important');
      (el as HTMLElement).style.setProperty('border-bottom', 'none', 'important');
    }
  });

  // AGGRESSIVE: Remove ALL borders from main page content containers
  // These often have subtle borders that appear as separator lines
  const pageContainers = document.querySelectorAll(`
    [class*="mainContent"],
    [class*="MainContent"],
    [class*="pageContent"],
    [class*="PageContent"],
    [class*="contentBox"],
    [class*="ContentBox"],
    [class*="articleRegion"],
    [class*="ArticleRegion"],
    [class*="bottomRegion"],
    [class*="BottomRegion"],
    main,
    article,
    #contentBox,
    #contentRow,
    [data-automation-id="contentScrollRegion"]
  `);
  pageContainers.forEach(el => {
    if (isInsidePanel(el)) {
      return;
    }
    if (!el.closest('[data-sp-web-part]')) {
      (el as HTMLElement).style.setProperty('border', 'none', 'important');
      (el as HTMLElement).style.setProperty('border-top', 'none', 'important');
      (el as HTMLElement).style.setProperty('border-bottom', 'none', 'important');
      (el as HTMLElement).style.setProperty('box-shadow', 'none', 'important');
    }
  });

  // Look for elements immediately after the last webpart that might be separators
  const lastWebpart = document.querySelector('[data-sp-web-part]:last-of-type');
  if (lastWebpart) {
    let sibling = lastWebpart.nextElementSibling;
    while (sibling) {
      if (isInsidePanel(sibling)) {
        sibling = sibling.nextElementSibling;
        continue;
      }

      // If it's not a Panel/Layer and not already marked, check if it looks like a separator
      if (!sibling.hasAttribute('data-jml-hidden')) {
        const siblingRect = sibling.getBoundingClientRect();
        const siblingStyle = window.getComputedStyle(sibling);

        // Hide if it's thin and has a visible background/border
        if (siblingRect.height > 0 && siblingRect.height <= 10) {
          const hasVisibleBg = siblingStyle.backgroundColor !== 'rgba(0, 0, 0, 0)' && siblingStyle.backgroundColor !== 'transparent';
          const hasBorder = siblingStyle.borderTopWidth !== '0px' || siblingStyle.borderBottomWidth !== '0px';

          if (hasVisibleBg || hasBorder) {
            (sibling as HTMLElement).style.setProperty('display', 'none', 'important');
            sibling.setAttribute('data-jml-hidden', 'true');
            console.log('[JML] Hidden post-webpart separator element');
          }
        }
      }
      sibling = sibling.nextElementSibling;
    }
  }
}

let spElementObserver: MutationObserver | null = null;

/**
 * Set up MutationObserver to catch dynamically loaded SharePoint elements
 */
function setupSharePointElementObserver(): void {
  if (spElementObserver) {
    return;
  }

  spElementObserver = new MutationObserver((mutations) => {
    let shouldCheck = false;

    mutations.forEach(mutation => {
      if (mutation.addedNodes.length > 0) {
        shouldCheck = true;
      }
    });

    if (shouldCheck) {
      // Debounce the check
      setTimeout(() => hideSharePointElements(), 100);
    }
  });

  spElementObserver.observe(document.body, {
    childList: true,
    subtree: true
  });

  // Also run periodically for the first few seconds to catch late-loading elements
  let checkCount = 0;
  const intervalId = setInterval(() => {
    hideSharePointElements();
    checkCount++;
    if (checkCount >= 10) {
      clearInterval(intervalId);
    }
  }, 500);
}

/**
 * Remove the portal styles observer.
 */
export function removePortalStyles(): void {
  if (observer) {
    observer.disconnect();
    observer = null;
  }

  document.body.removeAttribute(PORTAL_OBSERVER_ID);

  const fallbackStyle = document.getElementById('jml-portal-fallback-css');
  if (fallbackStyle) {
    fallbackStyle.remove();
  }

  // Remove data attributes from styled elements
  document.querySelectorAll('[data-jml-styled]').forEach(el => {
    el.removeAttribute('data-jml-styled');
  });

  isInitialized = false;
  console.log('[JML] Portal styles observer removed');
}

/**
 * React hook to inject portal styles on mount.
 */
export function usePortalStyles(): void {
  if (typeof window !== 'undefined') {
    injectPortalStyles();
  }
}

export default injectPortalStyles;
