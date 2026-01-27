// @ts-nocheck
// CacheService - In-memory caching for performance optimization
// Reduces SharePoint API calls by caching frequently accessed data

import { logger } from './LoggingService';

export interface ICacheEntry<T> {
  data: T;
  timestamp: number;
  expiresIn: number; // milliseconds
}

export class CacheService {
  private cache: Map<string, ICacheEntry<any>>;
  private static instance: CacheService;

  private constructor() {
    this.cache = new Map();
  }

  public static getInstance(): CacheService {
    if (!CacheService.instance) {
      CacheService.instance = new CacheService();
    }
    return CacheService.instance;
  }

  public get<T>(key: string): T | null {
    const entry = this.cache.get(key);
    if (!entry) {
      return null;
    }
    const now = Date.now();
    if (now - entry.timestamp > entry.expiresIn) {
      this.cache.delete(key);
      return null;
    }
    return entry.data as T;
  }

  public set<T>(key: string, data: T, expiresInMs: number = 300000): void {
    this.cache.set(key, {
      data,
      timestamp: Date.now(),
      expiresIn: expiresInMs
    });
  }

  public has(key: string): boolean {
    return this.get(key) !== null;
  }

  public remove(key: string): void {
    this.cache.delete(key);
  }

  public clear(): void {
    this.cache.clear();
  }

  public clearExpired(): void {
    const now = Date.now();
    const keysToDelete: string[] = [];
    this.cache.forEach((entry, key) => {
      if (now - entry.timestamp > entry.expiresIn) {
        keysToDelete.push(key);
      }
    });
    keysToDelete.forEach(key => this.cache.delete(key));
  }

  public async getOrFetch<T>(
    key: string,
    fetchFn: () => Promise<T>,
    expiresInMs: number = 300000
  ): Promise<T> {
    const cached = this.get<T>(key);
    if (cached !== null) {
      logger.debug('CacheService', `[Cache HIT] ${key}`);
      return cached;
    }
    logger.debug('CacheService', `[Cache MISS] ${key}`);
    const data = await fetchFn();
    this.set(key, data, expiresInMs);
    return data;
  }

  public invalidatePattern(pattern: string): void {
    const keysToDelete: string[] = [];
    this.cache.forEach((_, key) => {
      if (key.indexOf(pattern) !== -1) {
        keysToDelete.push(key);
      }
    });
    keysToDelete.forEach(key => this.cache.delete(key));
    logger.debug('CacheService', `[Cache] Invalidated ${keysToDelete.length} entries matching pattern: ${pattern}`);
  }

  public getStats(): { size: number; keys: string[] } {
    const keys: string[] = [];
    this.cache.forEach((_, key) => keys.push(key));
    return {
      size: this.cache.size,
      keys: keys
    };
  }
}

export class CacheKeys {
  static readonly PROCESSES_ALL = 'processes:all';
  static readonly PROCESS_BY_ID = (id: number) => `process:${id}`;
  static readonly TEMPLATES_ALL = 'templates:all';
  static readonly TEMPLATES_BY_TYPE = (type: string) => `templates:type:${type}`;
  static readonly TASKS_ALL = 'tasks:all';
  static readonly TASKS_BY_CATEGORY = (category: string) => `tasks:category:${category}`;
  static readonly TASK_ASSIGNMENTS_BY_PROCESS = (processId: number) => `taskassignments:process:${processId}`;
  static readonly MY_TASKS = (userId: number) => `mytasks:user:${userId}`;
  static readonly CONFIGURATIONS = 'configurations:all';
  static readonly CONFIG_VALUE = (key: string) => `config:${key}`;
  static readonly CURRENT_USER = 'currentuser';
}

export class CacheDurations {
  static readonly SHORT = 60000;
  static readonly MEDIUM = 300000;
  static readonly LONG = 900000;
  static readonly VERY_LONG = 1800000;
}
