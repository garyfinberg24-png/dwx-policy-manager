// @ts-nocheck
import { useCallback, useMemo } from 'react';
import {
  isEmbeddedMode,
  preserveEmbeddedMode,
  navigateWithEmbedded,
  exitEmbeddedMode,
  getExitEmbeddedUrl,
  buildJmlPageUrl,
  buildListUrl
} from '../utils/navigationUtils';

export interface IUseEmbeddedNavigationResult {
  /**
   * Whether the app is currently in embedded mode (SharePoint chrome hidden)
   */
  isEmbedded: boolean;

  /**
   * Navigate to a URL, preserving embedded mode
   */
  navigate: (url: string) => void;

  /**
   * Get a URL with embedded mode preserved
   */
  getUrl: (url: string) => string;

  /**
   * Exit embedded mode and show full SharePoint UI
   */
  exitToSharePoint: (url?: string) => void;

  /**
   * Get URL to exit embedded mode
   */
  getExitUrl: (url?: string) => string;

  /**
   * Build a JML page URL with embedded mode preserved
   */
  buildPageUrl: (
    siteUrl: string,
    pagePath: string,
    queryParams?: Record<string, string>
  ) => string;

  /**
   * Build a SharePoint list URL with embedded mode preserved
   */
  buildListUrl: (
    siteUrl: string,
    listName: string,
    view?: 'AllItems' | 'NewForm' | 'DispForm' | 'EditForm',
    itemId?: number,
    filterParams?: Record<string, string>
  ) => string;
}

/**
 * Hook for managing embedded navigation in JML application.
 *
 * Provides utilities for navigating between JML pages while preserving
 * the embedded mode (hidden SharePoint chrome) and allowing users to
 * exit back to full SharePoint view.
 *
 * @example
 * ```tsx
 * const { isEmbedded, navigate, exitToSharePoint } = useEmbeddedNavigation();
 *
 * // Navigate to another JML page (preserves embedded mode)
 * navigate('/sites/JML/SitePages/Dashboard.aspx');
 *
 * // Exit to full SharePoint view
 * exitToSharePoint();
 * ```
 */
export const useEmbeddedNavigation = (): IUseEmbeddedNavigationResult => {
  const isEmbedded = useMemo(() => isEmbeddedMode(), []);

  const navigate = useCallback((url: string) => {
    navigateWithEmbedded(url);
  }, []);

  const getUrl = useCallback((url: string) => {
    return preserveEmbeddedMode(url);
  }, []);

  const exitToSharePoint = useCallback((url?: string) => {
    exitEmbeddedMode(url);
  }, []);

  const getExitUrl = useCallback((url?: string) => {
    return getExitEmbeddedUrl(url);
  }, []);

  const buildPage = useCallback(
    (siteUrl: string, pagePath: string, queryParams?: Record<string, string>) => {
      return buildJmlPageUrl(siteUrl, pagePath, queryParams);
    },
    []
  );

  const buildList = useCallback(
    (
      siteUrl: string,
      listName: string,
      view?: 'AllItems' | 'NewForm' | 'DispForm' | 'EditForm',
      itemId?: number,
      filterParams?: Record<string, string>
    ) => {
      return buildListUrl(siteUrl, listName, view, itemId, filterParams);
    },
    []
  );

  return {
    isEmbedded,
    navigate,
    getUrl,
    exitToSharePoint,
    getExitUrl,
    buildPageUrl: buildPage,
    buildListUrl: buildList
  };
};

export default useEmbeddedNavigation;
