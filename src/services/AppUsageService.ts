// @ts-nocheck
import { UserRole } from './RoleDetectionService';
import { IAppDefinition, getAppsForRole } from '../webparts/jmlEmployeeDashboard/config/AppsConfig';

/**
 * Interface for tracking app usage data
 */
export interface IAppUsageRecord {
  /** Unique key for the app */
  appKey: string;
  /** Display name */
  appName: string;
  /** Fluent UI icon name */
  icon: string;
  /** Background color (hex) */
  color: string;
  /** Navigation URL */
  url: string;
  /** Number of times clicked */
  clickCount: number;
  /** Last used timestamp (ISO string) */
  lastUsed: string;
}

/**
 * Interface for stored user app usage data
 */
export interface IUserAppUsage {
  /** Schema version for migrations */
  version: number;
  /** User's role at time of storage */
  role: UserRole;
  /** Array of usage records */
  usage: IAppUsageRecord[];
  /** Last updated timestamp */
  lastUpdated: string;
}

/**
 * Configuration for app usage settings
 */
export interface IAppUsageConfig {
  /** Maximum number of apps to display (default: 8) */
  maxDisplayCount: number;
  /** Maximum apps to track in storage (default: 30) */
  maxTrackedCount: number;
  /** Days until usage decays completely (default: 30) */
  decayDays: number;
  /** Offset (gap) in pixels between app icons and system icons (default: 36) */
  appIconOffset: number;
}

/**
 * Default configuration
 */
const DEFAULT_CONFIG: IAppUsageConfig = {
  maxDisplayCount: 8,
  maxTrackedCount: 30,
  decayDays: 30,
  appIconOffset: 36
};

/**
 * Service for tracking and retrieving most-used apps
 * Uses localStorage for fast access with potential SharePoint sync
 */
export class AppUsageService {
  private static readonly STORAGE_KEY = 'jml-app-usage-v1';
  private static readonly CONFIG_KEY = 'jml-app-usage-config';

  private userRole: UserRole;
  private config: IAppUsageConfig;

  constructor(userRole: UserRole, config?: Partial<IAppUsageConfig>) {
    this.userRole = userRole;
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * Track when a user clicks/uses an app
   * @param app - The app definition or partial app data
   */
  public trackAppClick(app: IAppDefinition | { name: string; icon: string; color: string; url: string }): void {
    try {
      const usage = this.getUsageData();
      const appKey = app.name;
      const existingIndex = usage.findIndex(u => u.appKey === appKey);
      const now = new Date().toISOString();

      if (existingIndex >= 0) {
        // Update existing record
        usage[existingIndex].clickCount++;
        usage[existingIndex].lastUsed = now;
      } else {
        // Add new record
        usage.push({
          appKey: appKey,
          appName: app.name,
          icon: app.icon,
          color: app.color,
          url: app.url,
          clickCount: 1,
          lastUsed: now
        });
      }

      // Trim to max tracked count (keep highest scored)
      if (usage.length > this.config.maxTrackedCount) {
        const scored = usage.map(u => ({ ...u, score: this.calculateScore(u) }));
        scored.sort((a, b) => b.score - a.score);
        usage.length = 0;
        usage.push(...scored.slice(0, this.config.maxTrackedCount).map(({ score: _score, ...rest }) => rest));
      }

      this.saveUsageData(usage);
    } catch (error) {
      console.warn('Failed to track app usage:', error);
    }
  }

  /**
   * Get the most used apps for the current user
   * @param limit - Optional override for max apps to return
   * @returns Array of usage records sorted by score
   */
  public getMostUsedApps(limit?: number): IAppUsageRecord[] {
    const maxCount = limit ?? this.config.maxDisplayCount;
    const roleApps = getAppsForRole(this.userRole);
    const usage = this.getUsageData();

    // Filter to only apps available to this role
    const validUsage = usage.filter(u =>
      roleApps.some(app => app.name === u.appKey)
    );

    // Score and sort
    const scored = validUsage.map(u => ({
      ...u,
      score: this.calculateScore(u)
    }));

    return scored
      .sort((a, b) => b.score - a.score)
      .slice(0, maxCount)
      .map(({ score: _score, ...rest }) => rest);
  }

  /**
   * Get most used apps, filling with defaults if needed
   * Perfect for new users or users with limited activity
   * @param limit - Optional override for max apps to return
   * @returns Array of usage records (mix of actual and defaults)
   */
  public getMostUsedOrDefaults(limit?: number): IAppUsageRecord[] {
    const maxCount = limit ?? this.config.maxDisplayCount;
    const mostUsed = this.getMostUsedApps(maxCount);

    if (mostUsed.length >= maxCount) {
      return mostUsed;
    }

    // Fill remaining slots with role's default apps
    const roleApps = getAppsForRole(this.userRole);
    const usedKeys = new Set(mostUsed.map(u => u.appKey));

    const defaults = roleApps
      .filter(app => !usedKeys.has(app.name))
      .slice(0, maxCount - mostUsed.length)
      .map(app => ({
        appKey: app.name,
        appName: app.name,
        icon: app.icon,
        color: app.color,
        url: app.url,
        clickCount: 0,
        lastUsed: ''
      }));

    return [...mostUsed, ...defaults];
  }

  /**
   * Clear all usage data for the current user
   */
  public clearUsageData(): void {
    try {
      localStorage.removeItem(AppUsageService.STORAGE_KEY);
    } catch (error) {
      console.warn('Failed to clear usage data:', error);
    }
  }

  /**
   * Update the configuration (e.g., from Admin Panel)
   * @param config - Partial config to merge
   */
  public updateConfig(config: Partial<IAppUsageConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * Get current configuration
   */
  public getConfig(): IAppUsageConfig {
    return { ...this.config };
  }

  /**
   * Calculate a weighted score for an app based on frequency and recency
   * @param record - The usage record to score
   * @returns Score between 0 and 1
   */
  private calculateScore(record: IAppUsageRecord): number {
    const frequencyWeight = 0.6;
    const recencyWeight = 0.4;

    // Frequency score: normalized click count (caps at 50 clicks)
    const frequencyScore = Math.min(record.clickCount / 50, 1);

    // Recency score: decays over configured days
    let recencyScore = 0;
    if (record.lastUsed) {
      const daysSinceUse = (Date.now() - new Date(record.lastUsed).getTime()) / (1000 * 60 * 60 * 24);
      recencyScore = Math.max(0, 1 - (daysSinceUse / this.config.decayDays));
    }

    return (frequencyScore * frequencyWeight) + (recencyScore * recencyWeight);
  }

  /**
   * Get usage data from localStorage
   */
  private getUsageData(): IAppUsageRecord[] {
    try {
      const stored = localStorage.getItem(AppUsageService.STORAGE_KEY);
      if (!stored) return [];

      const data: IUserAppUsage = JSON.parse(stored);

      // Check if role changed - if so, return empty to rebuild
      if (data.role !== this.userRole) {
        return [];
      }

      return data.usage || [];
    } catch (error) {
      console.warn('Failed to load usage data:', error);
      return [];
    }
  }

  /**
   * Save usage data to localStorage
   */
  private saveUsageData(usage: IAppUsageRecord[]): void {
    try {
      const data: IUserAppUsage = {
        version: 1,
        role: this.userRole,
        usage: usage,
        lastUpdated: new Date().toISOString()
      };
      localStorage.setItem(AppUsageService.STORAGE_KEY, JSON.stringify(data));
    } catch (error) {
      console.warn('Failed to save usage data:', error);
    }
  }
}

/**
 * Singleton instance getter for the AppUsageService
 * Use this when you need quick access without creating new instances
 */
let serviceInstance: AppUsageService | null = null;

export function getAppUsageService(userRole: UserRole, config?: Partial<IAppUsageConfig>): AppUsageService {
  if (!serviceInstance || config) {
    serviceInstance = new AppUsageService(userRole, config);
  }
  return serviceInstance;
}

/**
 * Static helper to save config to localStorage (for Admin Panel)
 */
export function saveAppUsageConfig(config: Partial<IAppUsageConfig>): void {
  try {
    const existing = loadAppUsageConfig();
    const merged = { ...DEFAULT_CONFIG, ...existing, ...config };
    localStorage.setItem('jml-app-usage-config', JSON.stringify(merged));
  } catch (error) {
    console.warn('Failed to save app usage config:', error);
  }
}

/**
 * Static helper to load config from localStorage
 */
export function loadAppUsageConfig(): IAppUsageConfig {
  try {
    const stored = localStorage.getItem('jml-app-usage-config');
    if (stored) {
      return { ...DEFAULT_CONFIG, ...JSON.parse(stored) };
    }
  } catch (error) {
    console.warn('Failed to load app usage config:', error);
  }
  return DEFAULT_CONFIG;
}
