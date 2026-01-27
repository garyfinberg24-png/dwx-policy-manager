// @ts-nocheck
/**
 * Data Caching Hook
 * Implements SWR (stale-while-revalidate) pattern for SPFx data fetching
 */

import * as React from 'react';
import { logger } from '../services/LoggingService';

/**
 * Cache entry with metadata
 */
interface ICacheEntry<T> {
  data: T;
  timestamp: number;
  isStale: boolean;
}

/**
 * Configuration for data fetching
 */
export interface IUseDataCacheOptions<T> {
  /** Time in milliseconds before data is considered stale (default: 5 minutes) */
  staleTime?: number;
  /** Time in milliseconds before cached data expires completely (default: 30 minutes) */
  cacheTime?: number;
  /** Whether to revalidate on mount (default: true) */
  revalidateOnMount?: boolean;
  /** Whether to revalidate on window focus (default: false) */
  revalidateOnFocus?: boolean;
  /** Polling interval in milliseconds (0 = disabled) */
  pollingInterval?: number;
  /** Initial data before fetch completes */
  initialData?: T;
  /** Callback when fetch succeeds */
  onSuccess?: (data: T) => void;
  /** Callback when fetch fails */
  onError?: (error: Error) => void;
  /** Whether the query is enabled (default: true) */
  enabled?: boolean;
}

/**
 * Return type for useDataCache hook
 */
export interface IUseDataCacheResult<T> {
  /** The cached data */
  data: T | undefined;
  /** Loading state (first fetch only) */
  isLoading: boolean;
  /** Fetching state (any fetch including revalidation) */
  isFetching: boolean;
  /** Error if fetch failed */
  error: Error | undefined;
  /** Whether data is stale */
  isStale: boolean;
  /** Manually trigger revalidation */
  revalidate: () => Promise<void>;
  /** Mutate the cached data optimistically */
  mutate: (data: T | ((prev: T | undefined) => T)) => void;
}

// Global cache store
const globalCache = new Map<string, ICacheEntry<unknown>>();

// Active fetch promises to prevent duplicate requests
const activeFetches = new Map<string, Promise<unknown>>();

/**
 * Custom hook for data fetching with SWR-like caching
 * @param key Unique cache key
 * @param fetcher Async function to fetch data
 * @param options Configuration options
 */
export function useDataCache<T>(
  key: string | null,
  fetcher: () => Promise<T>,
  options: IUseDataCacheOptions<T> = {}
): IUseDataCacheResult<T> {
  const {
    staleTime = 5 * 60 * 1000, // 5 minutes
    cacheTime = 30 * 60 * 1000, // 30 minutes
    revalidateOnMount = true,
    revalidateOnFocus = false,
    pollingInterval = 0,
    initialData,
    onSuccess,
    onError,
    enabled = true
  } = options;

  const [state, setState] = React.useState<{
    data: T | undefined;
    isLoading: boolean;
    isFetching: boolean;
    error: Error | undefined;
    isStale: boolean;
  }>(() => {
    // Initialize from cache if available
    if (key) {
      const cached = globalCache.get(key) as ICacheEntry<T> | undefined;
      if (cached && Date.now() - cached.timestamp < cacheTime) {
        return {
          data: cached.data,
          isLoading: false,
          isFetching: false,
          error: undefined,
          isStale: Date.now() - cached.timestamp > staleTime
        };
      }
    }

    return {
      data: initialData,
      isLoading: !initialData,
      isFetching: false,
      error: undefined,
      isStale: true
    };
  });

  // Fetch function with deduplication
  const fetchData = React.useCallback(async () => {
    if (!key || !enabled) return;

    // Check if there's already an active fetch for this key
    const activeFetch = activeFetches.get(key);
    if (activeFetch) {
      return activeFetch as Promise<T>;
    }

    setState(prev => ({ ...prev, isFetching: true }));

    const fetchPromise = fetcher()
      .then((data) => {
        // Update cache
        globalCache.set(key, {
          data,
          timestamp: Date.now(),
          isStale: false
        });

        setState({
          data,
          isLoading: false,
          isFetching: false,
          error: undefined,
          isStale: false
        });

        if (onSuccess) {
          onSuccess(data);
        }

        logger.debug('useDataCache', `Data fetched and cached: ${key}`);
        return data;
      })
      .catch((error: Error) => {
        setState(prev => ({
          ...prev,
          isLoading: false,
          isFetching: false,
          error
        }));

        if (onError) {
          onError(error);
        }

        logger.error('useDataCache', `Failed to fetch data: ${key}`, error);
        throw error;
      })
      .finally(() => {
        activeFetches.delete(key);
      });

    activeFetches.set(key, fetchPromise);
    return fetchPromise;
  }, [key, fetcher, enabled, onSuccess, onError]);

  // Revalidate function exposed to consumers
  const revalidate = React.useCallback(async () => {
    await fetchData();
  }, [fetchData]);

  // Mutate function for optimistic updates
  const mutate = React.useCallback((newData: T | ((prev: T | undefined) => T)) => {
    if (!key) return;

    const resolvedData = typeof newData === 'function'
      ? (newData as (prev: T | undefined) => T)(state.data)
      : newData;

    // Update cache
    globalCache.set(key, {
      data: resolvedData,
      timestamp: Date.now(),
      isStale: false
    });

    setState(prev => ({
      ...prev,
      data: resolvedData,
      isStale: false
    }));

    logger.debug('useDataCache', `Data mutated: ${key}`);
  }, [key, state.data]);

  // Initial fetch on mount
  React.useEffect(() => {
    if (!key || !enabled) return;

    const cached = globalCache.get(key) as ICacheEntry<T> | undefined;
    const isExpired = !cached || Date.now() - cached.timestamp > cacheTime;
    const isStale = cached && Date.now() - cached.timestamp > staleTime;

    if (revalidateOnMount && (isExpired || isStale)) {
      fetchData();
    }
  }, [key, enabled, revalidateOnMount, staleTime, cacheTime, fetchData]);

  // Focus revalidation
  React.useEffect(() => {
    if (!revalidateOnFocus || !key || !enabled) return;

    const handleFocus = (): void => {
      const cached = globalCache.get(key) as ICacheEntry<T> | undefined;
      if (!cached || Date.now() - cached.timestamp > staleTime) {
        fetchData();
      }
    };

    window.addEventListener('focus', handleFocus);
    return () => window.removeEventListener('focus', handleFocus);
  }, [key, enabled, revalidateOnFocus, staleTime, fetchData]);

  // Polling
  React.useEffect(() => {
    if (!pollingInterval || pollingInterval <= 0 || !key || !enabled) return;

    const intervalId = setInterval(() => {
      fetchData();
    }, pollingInterval);

    return () => clearInterval(intervalId);
  }, [key, enabled, pollingInterval, fetchData]);

  return {
    data: state.data,
    isLoading: state.isLoading,
    isFetching: state.isFetching,
    error: state.error,
    isStale: state.isStale,
    revalidate,
    mutate
  };
}

/**
 * Invalidate cache entries by key pattern
 * @param pattern String or RegExp to match cache keys
 */
export function invalidateCache(pattern: string | RegExp): void {
  const keys = Array.from(globalCache.keys());

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const matches = typeof pattern === 'string'
      ? key.includes(pattern)
      : pattern.test(key);

    if (matches) {
      globalCache.delete(key);
      logger.debug('useDataCache', `Cache invalidated: ${key}`);
    }
  }
}

/**
 * Clear all cache entries
 */
export function clearCache(): void {
  globalCache.clear();
  logger.debug('useDataCache', 'All cache cleared');
}

/**
 * Get cache entry directly (for debugging/testing)
 */
export function getCacheEntry<T>(key: string): ICacheEntry<T> | undefined {
  return globalCache.get(key) as ICacheEntry<T> | undefined;
}

/**
 * Set cache entry directly (for hydration/prefetching)
 */
export function setCacheEntry<T>(key: string, data: T): void {
  globalCache.set(key, {
    data,
    timestamp: Date.now(),
    isStale: false
  });
}
