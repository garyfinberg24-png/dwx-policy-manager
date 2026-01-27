// useProcesses - React hook for process data

import { useState, useEffect, useCallback } from 'react';
import { IJmlProcess } from '../models';
import { SPService, CacheService, CacheKeys, CacheDurations } from '../services';

export interface IUseProcessesResult {
  processes: IJmlProcess[];
  loading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
  getProcessById: (id: number) => Promise<IJmlProcess | null>;
}

export function useProcesses(
  spService: SPService,
  filter?: string,
  orderBy?: string,
  top?: number
): IUseProcessesResult {
  const [processes, setProcesses] = useState<IJmlProcess[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const cache = CacheService.getInstance();

  const fetchProcesses = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const cacheKey = CacheKeys.PROCESSES_ALL;
      const data = await cache.getOrFetch(
        cacheKey,
        () => spService.getProcesses(filter, orderBy, top),
        CacheDurations.MEDIUM
      );

      setProcesses(data);
    } catch (err: any) {
      setError(err.message || 'Failed to fetch processes');
      console.error('Error in useProcesses:', err);
    } finally {
      setLoading(false);
    }
  }, [spService, filter, orderBy, top]);

  const refresh = useCallback(async () => {
    cache.invalidatePattern('process');
    await fetchProcesses();
  }, [fetchProcesses, cache]);

  const getProcessById = useCallback(async (id: number): Promise<IJmlProcess | null> => {
    try {
      const cacheKey = CacheKeys.PROCESS_BY_ID(id);
      const data = await cache.getOrFetch(
        cacheKey,
        () => spService.getProcessById(id),
        CacheDurations.MEDIUM
      );
      return data;
    } catch (err) {
      console.error(`Error fetching process ${id}:`, err);
      return null;
    }
  }, [spService, cache]);

  useEffect(() => {
    fetchProcesses();
  }, [fetchProcesses]);

  return {
    processes,
    loading,
    error,
    refresh,
    getProcessById
  };
}
