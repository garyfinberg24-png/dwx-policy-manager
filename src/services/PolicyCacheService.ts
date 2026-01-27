// @ts-nocheck
/**
 * Policy Cache Service
 * Provides in-memory caching with TTL for policy data
 * Reduces API calls and improves performance
 */

import { IPolicy, IPolicyAcknowledgement, IPolicyComplianceSummary } from '../models/IPolicy';
import { logger } from './LoggingService';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Cache entry with metadata
 */
interface ICacheEntry<T> {
  data: T;
  timestamp: number;
  ttl: number;
  key: string;
}

/**
 * Cache statistics
 */
export interface ICacheStats {
  hits: number;
  misses: number;
  size: number;
  hitRate: number;
}

/**
 * Paginated result
 */
export interface IPaginatedResult<T> {
  items: T[];
  totalCount: number;
  pageNumber: number;
  pageSize: number;
  totalPages: number;
  hasNextPage: boolean;
  hasPreviousPage: boolean;
}

/**
 * Cache configuration
 */
export interface ICacheConfig {
  defaultTtlMs: number;
  maxEntries: number;
  cleanupIntervalMs: number;
}

// ============================================================================
// DEFAULT CONFIGURATION
// ============================================================================

const DEFAULT_CONFIG: ICacheConfig = {
  defaultTtlMs: 5 * 60 * 1000, // 5 minutes
  maxEntries: 1000,
  cleanupIntervalMs: 60 * 1000 // 1 minute
};

// ============================================================================
// CACHE SERVICE
// ============================================================================

export class PolicyCacheService {
  private cache: Map<string, ICacheEntry<unknown>> = new Map();
  private config: ICacheConfig;
  private stats: ICacheStats = { hits: 0, misses: 0, size: 0, hitRate: 0 };
  private cleanupTimer: ReturnType<typeof setInterval> | null = null;

  // Cache key prefixes
  private static readonly CACHE_KEYS = {
    POLICY: 'policy:',
    POLICY_LIST: 'policy-list:',
    ACKNOWLEDGEMENT: 'ack:',
    ACK_LIST: 'ack-list:',
    COMPLIANCE: 'compliance:',
    USER_DASHBOARD: 'user-dashboard:',
    METRICS: 'metrics:',
    VERSIONS: 'versions:'
  };

  constructor(config?: Partial<ICacheConfig>) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.startCleanupTimer();
  }

  // ============================================================================
  // CORE CACHE OPERATIONS
  // ============================================================================

  /**
   * Get item from cache
   */
  public get<T>(key: string): T | null {
    const entry = this.cache.get(key) as ICacheEntry<T> | undefined;

    if (!entry) {
      this.stats.misses++;
      this.updateHitRate();
      return null;
    }

    // Check if expired
    if (this.isExpired(entry)) {
      this.cache.delete(key);
      this.stats.misses++;
      this.updateHitRate();
      return null;
    }

    this.stats.hits++;
    this.updateHitRate();
    return entry.data;
  }

  /**
   * Set item in cache
   */
  public set<T>(key: string, data: T, ttlMs?: number): void {
    // Evict oldest if at capacity
    if (this.cache.size >= this.config.maxEntries) {
      this.evictOldest();
    }

    const entry: ICacheEntry<T> = {
      key,
      data,
      timestamp: Date.now(),
      ttl: ttlMs ?? this.config.defaultTtlMs
    };

    this.cache.set(key, entry);
    this.stats.size = this.cache.size;
  }

  /**
   * Delete item from cache
   */
  public delete(key: string): boolean {
    const result = this.cache.delete(key);
    this.stats.size = this.cache.size;
    return result;
  }

  /**
   * Delete items matching a pattern
   */
  public deleteByPrefix(prefix: string): number {
    let count = 0;
    const keys = Array.from(this.cache.keys());
    keys.forEach(key => {
      if (key.startsWith(prefix)) {
        this.cache.delete(key);
        count++;
      }
    });
    this.stats.size = this.cache.size;
    return count;
  }

  /**
   * Clear entire cache
   */
  public clear(): void {
    this.cache.clear();
    this.stats = { hits: 0, misses: 0, size: 0, hitRate: 0 };
  }

  /**
   * Get cache statistics
   */
  public getStats(): ICacheStats {
    return { ...this.stats };
  }

  // ============================================================================
  // POLICY-SPECIFIC CACHE METHODS
  // ============================================================================

  /**
   * Get cached policy by ID
   */
  public getPolicy(policyId: number): IPolicy | null {
    return this.get<IPolicy>(`${PolicyCacheService.CACHE_KEYS.POLICY}${policyId}`);
  }

  /**
   * Cache a policy
   */
  public setPolicy(policy: IPolicy): void {
    if (policy.Id) {
      this.set(`${PolicyCacheService.CACHE_KEYS.POLICY}${policy.Id}`, policy);
    }
  }

  /**
   * Invalidate policy cache
   */
  public invalidatePolicy(policyId: number): void {
    this.delete(`${PolicyCacheService.CACHE_KEYS.POLICY}${policyId}`);
    // Also invalidate any list caches that might contain this policy
    this.deleteByPrefix(PolicyCacheService.CACHE_KEYS.POLICY_LIST);
    this.deleteByPrefix(PolicyCacheService.CACHE_KEYS.COMPLIANCE);
    this.deleteByPrefix(PolicyCacheService.CACHE_KEYS.METRICS);
  }

  /**
   * Get cached policy list
   */
  public getPolicyList(cacheKey: string): IPolicy[] | null {
    return this.get<IPolicy[]>(`${PolicyCacheService.CACHE_KEYS.POLICY_LIST}${cacheKey}`);
  }

  /**
   * Cache policy list
   */
  public setPolicyList(cacheKey: string, policies: IPolicy[], ttlMs?: number): void {
    this.set(`${PolicyCacheService.CACHE_KEYS.POLICY_LIST}${cacheKey}`, policies, ttlMs);
    // Also cache individual policies
    policies.forEach(policy => this.setPolicy(policy));
  }

  /**
   * Get cached acknowledgement
   */
  public getAcknowledgement(ackId: number): IPolicyAcknowledgement | null {
    return this.get<IPolicyAcknowledgement>(`${PolicyCacheService.CACHE_KEYS.ACKNOWLEDGEMENT}${ackId}`);
  }

  /**
   * Cache acknowledgement
   */
  public setAcknowledgement(ack: IPolicyAcknowledgement): void {
    if (ack.Id) {
      this.set(`${PolicyCacheService.CACHE_KEYS.ACKNOWLEDGEMENT}${ack.Id}`, ack);
    }
  }

  /**
   * Invalidate acknowledgement cache
   */
  public invalidateAcknowledgement(ackId: number): void {
    this.delete(`${PolicyCacheService.CACHE_KEYS.ACKNOWLEDGEMENT}${ackId}`);
    this.deleteByPrefix(PolicyCacheService.CACHE_KEYS.ACK_LIST);
    this.deleteByPrefix(PolicyCacheService.CACHE_KEYS.USER_DASHBOARD);
    this.deleteByPrefix(PolicyCacheService.CACHE_KEYS.COMPLIANCE);
  }

  /**
   * Get cached compliance summary
   */
  public getComplianceSummary(policyId: number): IPolicyComplianceSummary | null {
    return this.get<IPolicyComplianceSummary>(`${PolicyCacheService.CACHE_KEYS.COMPLIANCE}${policyId}`);
  }

  /**
   * Cache compliance summary
   */
  public setComplianceSummary(policyId: number, summary: IPolicyComplianceSummary): void {
    this.set(`${PolicyCacheService.CACHE_KEYS.COMPLIANCE}${policyId}`, summary);
  }

  /**
   * Get cached user dashboard
   */
  public getUserDashboard(userId: number): unknown | null {
    return this.get(`${PolicyCacheService.CACHE_KEYS.USER_DASHBOARD}${userId}`);
  }

  /**
   * Cache user dashboard
   */
  public setUserDashboard(userId: number, dashboard: unknown): void {
    // Shorter TTL for user dashboards as they change frequently
    this.set(`${PolicyCacheService.CACHE_KEYS.USER_DASHBOARD}${userId}`, dashboard, 2 * 60 * 1000);
  }

  /**
   * Invalidate user dashboard cache
   */
  public invalidateUserDashboard(userId: number): void {
    this.delete(`${PolicyCacheService.CACHE_KEYS.USER_DASHBOARD}${userId}`);
  }

  /**
   * Get cached policy versions
   */
  public getPolicyVersions(policyId: number): unknown[] | null {
    return this.get<unknown[]>(`${PolicyCacheService.CACHE_KEYS.VERSIONS}${policyId}`);
  }

  /**
   * Cache policy versions
   */
  public setPolicyVersions(policyId: number, versions: unknown[]): void {
    this.set(`${PolicyCacheService.CACHE_KEYS.VERSIONS}${policyId}`, versions);
  }

  // ============================================================================
  // PRIVATE METHODS
  // ============================================================================

  /**
   * Check if cache entry is expired
   */
  private isExpired(entry: ICacheEntry<unknown>): boolean {
    return Date.now() - entry.timestamp > entry.ttl;
  }

  /**
   * Evict oldest cache entry
   */
  private evictOldest(): void {
    let oldestKey: string | null = null;
    let oldestTime = Infinity;

    const entries = Array.from(this.cache.entries());
    entries.forEach(([key, entry]) => {
      if (entry.timestamp < oldestTime) {
        oldestTime = entry.timestamp;
        oldestKey = key;
      }
    });

    if (oldestKey) {
      this.cache.delete(oldestKey);
      logger.info('PolicyCacheService', `Evicted oldest cache entry: ${oldestKey}`);
    }
  }

  /**
   * Update hit rate statistic
   */
  private updateHitRate(): void {
    const total = this.stats.hits + this.stats.misses;
    this.stats.hitRate = total > 0 ? (this.stats.hits / total) * 100 : 0;
  }

  /**
   * Start cleanup timer
   */
  private startCleanupTimer(): void {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
    }

    this.cleanupTimer = setInterval(() => {
      this.cleanup();
    }, this.config.cleanupIntervalMs);
  }

  /**
   * Clean up expired entries
   */
  private cleanup(): void {
    let cleaned = 0;
    const entries = Array.from(this.cache.entries());
    entries.forEach(([key, entry]) => {
      if (this.isExpired(entry)) {
        this.cache.delete(key);
        cleaned++;
      }
    });

    if (cleaned > 0) {
      this.stats.size = this.cache.size;
      logger.info('PolicyCacheService', `Cleaned ${cleaned} expired cache entries`);
    }
  }

  /**
   * Stop cleanup timer (for disposal)
   */
  public dispose(): void {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = null;
    }
    this.clear();
  }
}

// ============================================================================
// SINGLETON INSTANCE
// ============================================================================

let cacheInstance: PolicyCacheService | null = null;

/**
 * Get the singleton cache instance
 */
export function getPolicyCacheService(config?: Partial<ICacheConfig>): PolicyCacheService {
  if (!cacheInstance) {
    cacheInstance = new PolicyCacheService(config);
  }
  return cacheInstance;
}

/**
 * Reset the cache instance (for testing)
 */
export function resetPolicyCacheService(): void {
  if (cacheInstance) {
    cacheInstance.dispose();
    cacheInstance = null;
  }
}

// ============================================================================
// PAGINATION UTILITIES
// ============================================================================

/**
 * Paginate an array of items
 */
export function paginateArray<T>(
  items: T[],
  pageNumber: number,
  pageSize: number
): IPaginatedResult<T> {
  const totalCount = items.length;
  const totalPages = Math.ceil(totalCount / pageSize);
  const startIndex = (pageNumber - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedItems = items.slice(startIndex, endIndex);

  return {
    items: paginatedItems,
    totalCount,
    pageNumber,
    pageSize,
    totalPages,
    hasNextPage: pageNumber < totalPages,
    hasPreviousPage: pageNumber > 1
  };
}

/**
 * Generate cache key from filters
 */
export function generateCacheKey(filters: Record<string, unknown>): string {
  const sortedKeys = Object.keys(filters).sort();
  const parts = sortedKeys.map(key => {
    const value = filters[key];
    if (value === undefined || value === null) return '';
    return `${key}:${JSON.stringify(value)}`;
  }).filter(Boolean);

  return parts.join('|') || 'all';
}
