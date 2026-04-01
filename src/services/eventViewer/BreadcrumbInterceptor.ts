/**
 * BreadcrumbInterceptor — Captures user interactions (clicks, navigation,
 * hash changes) as breadcrumbs for error replay context.
 *
 * Maintains a ring buffer of recent interactions. When an error occurs,
 * the breadcrumbs show what the user was doing leading up to the error.
 *
 * Usage:
 *   BreadcrumbInterceptor.getInstance().install();
 *   // ... all clicks + navigations are captured ...
 *   BreadcrumbInterceptor.getInstance().uninstall();
 *   const crumbs = BreadcrumbInterceptor.getInstance().getBreadcrumbs();
 */

import { IUIBreadcrumb } from '../../models/IEventViewer';

const MAX_BREADCRUMBS = 50;

export class BreadcrumbInterceptor {
  private static _instance: BreadcrumbInterceptor;

  private _installed: boolean = false;
  private _breadcrumbs: IUIBreadcrumb[] = [];
  private _clickHandler: ((e: MouseEvent) => void) | null = null;
  private _hashHandler: (() => void) | null = null;
  private _popstateHandler: (() => void) | null = null;

  private constructor() {}

  public static getInstance(): BreadcrumbInterceptor {
    if (!BreadcrumbInterceptor._instance) {
      BreadcrumbInterceptor._instance = new BreadcrumbInterceptor();
    }
    return BreadcrumbInterceptor._instance;
  }

  // ==========================================================================
  // INSTALL / UNINSTALL
  // ==========================================================================

  public install(): void {
    if (this._installed || typeof window === 'undefined') return;

    // Capture clicks
    this._clickHandler = (e: MouseEvent) => {
      try {
        const target = e.target as HTMLElement;
        if (!target) return;

        const tag = target.tagName?.toLowerCase() || '';
        const text = (target.textContent || '').trim().substring(0, 60);
        const ariaLabel = target.getAttribute('aria-label') || '';
        const role = target.getAttribute('role') || '';
        const className = target.className ? (typeof target.className === 'string' ? target.className.split(' ')[0] : '') : '';

        // Build description
        let desc = '';
        if (tag === 'button' || role === 'button') {
          desc = `Clicked button: "${text || ariaLabel || className}"`;
        } else if (tag === 'a') {
          desc = `Clicked link: "${text || (target as HTMLAnchorElement).href || ''}"`;
        } else if (tag === 'input' || tag === 'select' || tag === 'textarea') {
          desc = `Focused ${tag}: ${ariaLabel || target.getAttribute('name') || className}`;
        } else if (text && text.length > 0) {
          desc = `Clicked: "${text.substring(0, 40)}"`;
        } else {
          // Skip generic div/span clicks with no meaningful text
          return;
        }

        // Build selector path
        const selector = this._buildSelector(target);

        this._push({
          timestamp: new Date().toISOString(),
          type: 'click',
          description: desc,
          target: selector,
          pageUrl: window.location.pathname + window.location.hash,
        });
      } catch (_) {
        // Never break the app on breadcrumb failure
      }
    };

    // Capture hash changes (SPA navigation)
    this._hashHandler = () => {
      this._push({
        timestamp: new Date().toISOString(),
        type: 'navigation',
        description: `Navigated to ${window.location.pathname}${window.location.hash}`,
        pageUrl: window.location.pathname + window.location.hash,
      });
    };

    // Capture popstate (browser back/forward)
    this._popstateHandler = () => {
      this._push({
        timestamp: new Date().toISOString(),
        type: 'navigation',
        description: `Browser back/forward to ${window.location.pathname}`,
        pageUrl: window.location.pathname + window.location.hash,
      });
    };

    document.addEventListener('click', this._clickHandler, true);
    window.addEventListener('hashchange', this._hashHandler);
    window.addEventListener('popstate', this._popstateHandler);

    this._installed = true;
  }

  public uninstall(): void {
    if (!this._installed) return;

    if (this._clickHandler) {
      document.removeEventListener('click', this._clickHandler, true);
      this._clickHandler = null;
    }
    if (this._hashHandler) {
      window.removeEventListener('hashchange', this._hashHandler);
      this._hashHandler = null;
    }
    if (this._popstateHandler) {
      window.removeEventListener('popstate', this._popstateHandler);
      this._popstateHandler = null;
    }

    this._installed = false;
  }

  public get isInstalled(): boolean {
    return this._installed;
  }

  // ==========================================================================
  // BREADCRUMB ACCESS
  // ==========================================================================

  /** Get all breadcrumbs (newest last) */
  public getBreadcrumbs(): IUIBreadcrumb[] {
    return this._breadcrumbs.slice();
  }

  /** Get the most recent N breadcrumbs */
  public getRecent(count: number): IUIBreadcrumb[] {
    return this._breadcrumbs.slice(-count);
  }

  /** Add a custom breadcrumb programmatically */
  public addCustom(description: string): void {
    this._push({
      timestamp: new Date().toISOString(),
      type: 'custom',
      description,
      pageUrl: typeof window !== 'undefined' ? window.location.pathname : '',
    });
  }

  /** Clear all breadcrumbs */
  public clear(): void {
    this._breadcrumbs = [];
  }

  // ==========================================================================
  // INTERNAL
  // ==========================================================================

  private _push(crumb: IUIBreadcrumb): void {
    this._breadcrumbs.push(crumb);
    if (this._breadcrumbs.length > MAX_BREADCRUMBS) {
      this._breadcrumbs.shift();
    }
  }

  /** Build a short CSS-like selector for the target element */
  private _buildSelector(el: HTMLElement): string {
    try {
      const parts: string[] = [];
      let current: HTMLElement | null = el;
      let depth = 0;

      while (current && depth < 3) {
        const tag = current.tagName?.toLowerCase() || '';
        if (!tag || tag === 'html' || tag === 'body') break;

        let sel = tag;
        if (current.id) {
          sel += '#' + current.id;
        } else if (current.className && typeof current.className === 'string') {
          const cls = current.className.trim().split(/\s+/)[0];
          if (cls && cls.length < 40) sel += '.' + cls;
        }
        parts.unshift(sel);
        current = current.parentElement;
        depth++;
      }

      return parts.join(' > ');
    } catch (_) {
      return '';
    }
  }
}
