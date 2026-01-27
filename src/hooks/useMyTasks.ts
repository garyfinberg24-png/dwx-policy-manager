// useMyTasks - React hook for current user's tasks

import { useState, useEffect, useCallback } from 'react';
import { IJmlTaskAssignment } from '../models';
import { SPService, CacheService, CacheKeys, CacheDurations } from '../services';

export interface IUseMyTasksResult {
  tasks: IJmlTaskAssignment[];
  loading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
  updateTask: (id: number, updates: Partial<IJmlTaskAssignment>) => Promise<void>;
}

export function useMyTasks(
  spService: SPService,
  userId: number
): IUseMyTasksResult {
  const [tasks, setTasks] = useState<IJmlTaskAssignment[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const cache = CacheService.getInstance();

  const fetchTasks = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const cacheKey = CacheKeys.MY_TASKS(userId);
      const data = await cache.getOrFetch(
        cacheKey,
        () => spService.getMyTasks(userId),
        CacheDurations.SHORT // Shorter cache for my tasks
      );

      setTasks(data);
    } catch (err: any) {
      setError(err.message || 'Failed to fetch tasks');
      console.error('Error in useMyTasks:', err);
    } finally {
      setLoading(false);
    }
  }, [spService, userId, cache]);

  const refresh = useCallback(async () => {
    cache.remove(CacheKeys.MY_TASKS(userId));
    await fetchTasks();
  }, [fetchTasks, cache, userId]);

  const updateTask = useCallback(async (id: number, updates: Partial<IJmlTaskAssignment>) => {
    try {
      await spService.updateTaskAssignment(id, updates);
      cache.remove(CacheKeys.MY_TASKS(userId));
      cache.invalidatePattern('taskassignments');
      await fetchTasks();
    } catch (err: any) {
      setError(err.message || 'Failed to update task');
      throw err;
    }
  }, [spService, userId, cache, fetchTasks]);

  useEffect(() => {
    if (userId) {
      fetchTasks();
    }
  }, [fetchTasks, userId]);

  return {
    tasks,
    loading,
    error,
    refresh,
    updateTask
  };
}
