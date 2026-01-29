// @ts-nocheck
/* eslint-disable */
/**
 * SharePoint Global Overrides Utility
 *
 * This utility injects CSS styles to enable full-bleed layouts and handle
 * embedded mode (?env=embedded) for kiosk/display scenarios.
 *
 * Call injectSharePointOverrides() from any webpart's onInit() to ensure
 * consistent behavior across all JML webparts.
 *
 * Also supports injecting the JML App Shell (header + footer) via:
 * - injectSharePointOverrides(context, { injectAppShell: true })
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { injectJmlAppShell, IAppShellOptions } from './JmlAppShellInjector';

const PM_STYLES_ID = 'jml-global-overrides';
const PM_EMBEDDED_STYLES_ID = 'jml-embedded-mode';

/**
 * Options for injectSharePointOverrides
 */
export interface ISharePointOverridesOptions {
  /** Whether to inject the JML App Shell (header + footer) */
  injectAppShell?: boolean;
  /** Options to pass to the App Shell injector */
  appShellOptions?: IAppShellOptions;
}

/**
 * Global CSS styles for full-bleed layouts and social bar hiding.
 */
const PM_GLOBAL_STYLES = `
/* =============================================================================
   JML GLOBAL OVERRIDES - Injected by SharePointOverrides utility
   ============================================================================= */

/* SHAREPOINT CANVAS OVERRIDES - Enable full-bleed layouts */
.CanvasZone,
[class*="CanvasZone"],
.CanvasSection,
[class*="CanvasSection"],
.ControlZone,
[class*="ControlZone"],
.CanvasZoneContainer,
[class*="CanvasZoneContainer"] {
  overflow: visible !important;
  max-width: none !important;
}

[data-sp-web-part],
.webPartContainer,
[class*="webPartContainer"] {
  overflow: visible !important;
}

[class*="CanvasComponent"],
.CanvasComponent {
  overflow: visible !important;
}

/* HIDE SHAREPOINT SOCIAL BAR (Like, Views, Save for Later, Comments) */
/* IMPORTANT: Exclude Panel/Layer/Overlay elements to prevent breaking fly-in panels */
[class*="pageReactions"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="PageReactions"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="pageLikes"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="PageLikes"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[data-automation-id="PageLikes"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="socialBar"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="SocialBar"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="pageFooter"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="PageFooter"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[data-automation-id="pageFooter"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[data-automation-id="pageComments"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
.ms-CommentsWrapper:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="CommentsWrapper"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="pageActions"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *):not(.ms-Panel-actions),
[class*="PageActions"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *):not(.ms-Panel-actions),
[class*="saveForLater"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="SaveForLater"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[data-automation-id="saveForLater"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="viewCount"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="ViewCount"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[data-automation-id="viewCount"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="pageMetadata"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
[class*="PageMetadata"]:not(.ms-Panel *):not(.ms-Layer *):not(.ms-Overlay *),
div[class*="root_"][class*="pageFooter"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="root_"][class*="socialBar"]:not(.ms-Panel *):not(.ms-Layer *),
/* Additional selectors for SharePoint Modern social elements */
[data-automation-id="pageSocialBar"]:not(.ms-Panel *):not(.ms-Layer *),
[data-automation-id="pageLikes"]:not(.ms-Panel *):not(.ms-Layer *),
[data-sp-feature-tag="PageSocialBar"]:not(.ms-Panel *):not(.ms-Layer *),
[data-sp-feature-tag="PageReactions"]:not(.ms-Panel *):not(.ms-Layer *),
footer[class*="pageFooter"]:not(.ms-Panel *):not(.ms-Layer *),
section[class*="pageReactions"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="pageLikes_"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="socialContainer"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="SocialContainer"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="reactionBar"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="ReactionBar"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="pageInteractions"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="PageInteractions"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="likesAndComments"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="LikesAndComments"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="pageSocialSection"]:not(.ms-Panel *):not(.ms-Layer *),
div[class*="PageSocialSection"]:not(.ms-Panel *):not(.ms-Layer *),
/* Hide footer under webpart zone */
.CanvasZone + [class*="pageFooter"]:not(.ms-Panel *):not(.ms-Layer *),
.CanvasZone ~ [class*="socialBar"]:not(.ms-Panel *):not(.ms-Layer *),
.CanvasZone ~ footer:not(.ms-Panel *):not(.ms-Layer *),
[class*="CanvasZone"] + footer:not(.ms-Panel *):not(.ms-Layer *),
[class*="mainContent"] > footer:not(.ms-Panel *):not(.ms-Layer *),
#spPageCanvasContent + footer:not(.ms-Panel *):not(.ms-Layer *) {
  display: none !important;
}

/* AGGRESSIVE: Hide ALL footers except JML footers */
footer:not([data-sp-web-part] *):not(.ms-Panel *):not(.ms-Layer *):not([class*="jml"]):not([data-jml-footer]) {
  display: none !important;
  visibility: hidden !important;
  height: 0 !important;
}

/* Hide separator lines / HR elements */
hr:not([data-sp-web-part] *):not(.ms-Panel *):not(.ms-Layer *) {
  display: none !important;
  height: 0 !important;
  border: none !important;
}

/* Remove borders from content containers that could look like separators */
[class*="mainContent"]:not(.ms-Panel *),
[class*="pageContent"]:not(.ms-Panel *),
[class*="contentBox"]:not(.ms-Panel *),
main:not(.ms-Panel *),
article:not(.ms-Panel *) {
  border: none !important;
  border-top: none !important;
  border-bottom: none !important;
}

/* Hide buttons by aria-label (Like, Save for later) */
button[aria-label="Like"]:not(.ms-Panel *),
button[aria-label="Save for later"]:not(.ms-Panel *) {
  display: none !important;
}

/* Hide social bar containers */
[class*="likesAndComments"]:not(.ms-Panel *),
[class*="LikesAndComments"]:not(.ms-Panel *),
[class*="pageSocialSection"]:not(.ms-Panel *),
[class*="socialContainer"]:not(.ms-Panel *),
[class*="SocialContainer"]:not(.ms-Panel *) {
  display: none !important;
}
`;

/**
 * CSS styles for embedded mode - hides all SharePoint chrome
 */
const PM_EMBEDDED_STYLES = `
/* EMBEDDED MODE - Hide all SharePoint chrome */
#SuiteNavPlaceHolder,
[class*="SuiteNav"],
.ms-HubNav,
[class*="HubNav"],
#spSiteHeader,
[data-automationid="SiteHeader"],
[class*="siteHeader"],
[class*="SiteHeader"],
.ms-siteHeader-container,
#spCommandBar,
[class*="commandBar"],
[class*="CommandBar"],
.sp-appBar,
[class*="appBar"],
#sp-appBar,
[class*="spAppBar"],
.ms-FocusZone[role="navigation"],
[data-automationid="pageHeader"],
/* NOTE: Do NOT use [class*="pageHeader"] - it hides our JmlAppHeader pageHeader section! */
/* Only target SharePoint's specific page header classes */
.ms-compositeHeader,
[class*="compositeHeader"],
[class*="titleRow"],
[class*="TitleRow"],
#SuiteNavWrapper,
.o365cs-nav-container,
[class*="o365cs-nav"],
.od-TopBar,
[class*="TopBar"],
.ms-siteHeader,
div[class*="titleRegion"],
div[class*="TitleRegion"],
#spLeftNav,
[class*="leftNav"],
[class*="LeftNav"],
.ms-Nav,
[data-automationid="VerticalNav"] {
  display: none !important;
}

/* Remove padding/margins that account for hidden elements */
#workbenchPageContent,
[class*="workbenchPageContent"],
.SPCanvas,
[class*="SPCanvas"],
.CanvasZone,
[class*="mainContent"] {
  margin-top: 0 !important;
  padding-top: 0 !important;
}

/* Ensure body starts at top */
body {
  padding-top: 0 !important;
  margin-top: 0 !important;
}

/* Full viewport height for embedded content */
.CanvasZone,
[class*="CanvasZone"] {
  min-height: 100vh !important;
}
`;

/**
 * Checks if the page should show full SharePoint chrome.
 * By default, JML hides SharePoint chrome (app-like experience).
 * Use ?env=full to show SharePoint chrome if needed.
 */
function shouldShowSharePointChrome(): boolean {
  if (typeof window === 'undefined') return false;
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get('env') === 'full';
}

/**
 * Injects global SharePoint override styles into the page.
 * Safe to call multiple times - only injects once.
 *
 * By default, this applies BOTH:
 * - Full-bleed layout (canvas overflow fixes)
 * - Embedded mode styling (hides SharePoint chrome for app-like experience)
 *
 * To show SharePoint chrome, append ?env=full to the URL.
 *
 * @param context - Optional WebPartContext for user info in App Shell
 * @param options - Optional configuration options
 * @returns void
 */
export function injectSharePointOverrides(
  context?: WebPartContext,
  options?: ISharePointOverridesOptions
): void {
  // Skip if running in Node.js (SSR) or already injected
  if (typeof document === 'undefined') return;

  // Inject global overrides (full-bleed + social bar hiding)
  if (!document.getElementById(PM_STYLES_ID)) {
    const styleElement = document.createElement('style');
    styleElement.id = PM_STYLES_ID;
    styleElement.type = 'text/css';
    styleElement.textContent = PM_GLOBAL_STYLES;
    document.head.appendChild(styleElement);
    console.log('[JML] Injected global SharePoint overrides');
  }

  // ALWAYS inject embedded mode styles UNLESS ?env=full is specified
  // This gives an app-like experience by default
  if (!shouldShowSharePointChrome() && !document.getElementById(PM_EMBEDDED_STYLES_ID)) {
    const embeddedStyleElement = document.createElement('style');
    embeddedStyleElement.id = PM_EMBEDDED_STYLES_ID;
    embeddedStyleElement.type = 'text/css';
    embeddedStyleElement.textContent = PM_EMBEDDED_STYLES;
    document.head.appendChild(embeddedStyleElement);
    console.log('[JML] Injected app-mode styles (hiding SharePoint chrome)');
  }

  // Optionally inject the JML App Shell (header + footer)
  if (options?.injectAppShell) {
    injectJmlAppShell(context, options.appShellOptions);
    console.log('[JML] Injected JML App Shell (header + footer)');
  }
}

/**
 * Removes the injected SharePoint override styles and App Shell.
 * Useful for cleanup in testing scenarios.
 */
export function removeSharePointOverrides(): void {
  if (typeof document === 'undefined') return;

  const globalStyles = document.getElementById(PM_STYLES_ID);
  if (globalStyles) {
    globalStyles.remove();
  }

  const embeddedStyles = document.getElementById(PM_EMBEDDED_STYLES_ID);
  if (embeddedStyles) {
    embeddedStyles.remove();
  }

  // Also remove App Shell if it exists
  const { removeJmlAppShell } = require('./JmlAppShellInjector');
  removeJmlAppShell();
}

// Note: isEmbeddedMode is already exported from navigationUtils
// Use shouldShowSharePointChrome for internal logic here

// =============================================================================
// CRITICAL: IMMEDIATE CSS INJECTION TO PREVENT FOUC
// This runs at module load time (before React mounts) to hide SP chrome ASAP
// =============================================================================
(function injectCriticalCssImmediately(): void {
  // Skip if running in Node.js (SSR)
  if (typeof document === 'undefined' || typeof window === 'undefined') return;

  // Check if we should show SharePoint chrome
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('env') === 'full') return;

  // Critical CSS ID - different from main styles to avoid conflicts
  const CRITICAL_CSS_ID = 'jml-critical-fouc-fix';
  if (document.getElementById(CRITICAL_CSS_ID)) return;

  // Critical CSS to hide SharePoint chrome IMMEDIATELY
  // This is a minimal subset focused on preventing the flash
  const criticalCss = `
    /* CRITICAL CSS - Injected immediately to prevent FOUC */
    /* Hide SharePoint navigation/header elements */
    #SuiteNavPlaceHolder,
    [class*="SuiteNav"],
    .ms-HubNav,
    [class*="HubNav"],
    #spSiteHeader,
    [data-automationid="SiteHeader"],
    [class*="siteHeader"],
    [class*="SiteHeader"],
    .ms-siteHeader-container,
    #spCommandBar,
    .sp-appBar,
    [class*="appBar"],
    #sp-appBar,
    .ms-FocusZone[role="navigation"],
    [data-automationid="pageHeader"],
    .ms-compositeHeader,
    [class*="compositeHeader"],
    #SuiteNavWrapper,
    .o365cs-nav-container,
    [class*="o365cs-nav"],
    .od-TopBar,
    [class*="TopBar"],
    .ms-siteHeader,
    #spLeftNav,
    [class*="leftNav"],
    [class*="LeftNav"],
    .ms-Nav,
    [data-automationid="VerticalNav"] {
      display: none !important;
      visibility: hidden !important;
    }

    /* Remove top padding/margins that account for hidden elements */
    #workbenchPageContent,
    [class*="workbenchPageContent"],
    .SPCanvas,
    [class*="SPCanvas"],
    .CanvasZone,
    [class*="mainContent"] {
      margin-top: 0 !important;
      padding-top: 0 !important;
    }

    body {
      padding-top: 0 !important;
      margin-top: 0 !important;
    }
  `;

  // Inject the critical CSS as early as possible
  const styleElement = document.createElement('style');
  styleElement.id = CRITICAL_CSS_ID;
  styleElement.type = 'text/css';
  styleElement.textContent = criticalCss;

  // Insert at the beginning of head for highest priority
  if (document.head.firstChild) {
    document.head.insertBefore(styleElement, document.head.firstChild);
  } else {
    document.head.appendChild(styleElement);
  }

  console.log('[JML] Critical FOUC-prevention CSS injected immediately');
})();
