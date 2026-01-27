// @ts-nocheck
// UserPreferencesService - Manages user preferences and personalization
// Handles dashboard layouts, themes, notifications, favorites, and localization

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';
import { logger } from './LoggingService';
import {
  IUserPreferences,
  IDashboardLayout,
  ISavedFilter,
  ISavedView,
  INotificationSettings,
  IFavoriteItem,
  DEFAULT_PREFERENCES,
  ICustomTheme,
  ActivityType
} from '../models/IUserPreferences';
import { ValidationUtils } from '../utils/ValidationUtils';

export class UserPreferencesService {
  private sp: SPFI;
  private currentUserId: string = '';
  private currentUserEmail: string = '';
  private cachedPreferences: IUserPreferences | null = null;
  private cacheExpiry: number = 5 * 60 * 1000; // 5 minutes
  private lastCacheTime: number = 0;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize service with current user info
   */
  public async initialize(): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();
      this.currentUserId = currentUser.Id.toString();
      this.currentUserEmail = currentUser.Email;
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to initialize UserPreferencesService:', error);
    }
  }

  /**
   * Get user preferences
   */
  public async getPreferences(userId?: string): Promise<IUserPreferences> {
    const targetUserId = userId || this.currentUserId;

    // Validate user ID
    ValidationUtils.validateUserId(targetUserId);

    // Check cache first
    if (this.cachedPreferences &&
        this.cachedPreferences.UserId === targetUserId &&
        Date.now() - this.lastCacheTime < this.cacheExpiry) {
      return this.cachedPreferences;
    }

    try {
      // Try to get existing preferences with secure filter
      const filter = ValidationUtils.buildFilter('UserId', 'eq', targetUserId);
      const items = await this.sp.web.lists
        .getByTitle('JML_UserPreferences')
        .items
        .filter(filter)
        .top(1)();

      if (items.length > 0) {
        const prefs = this.parsePreferences(items[0]);
        this.cachePreferences(prefs);
        return prefs;
      } else {
        // Create default preferences for new user
        const defaultPrefs = await this.createDefaultPreferences(targetUserId);
        this.cachePreferences(defaultPrefs);
        return defaultPrefs;
      }
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to get user preferences:', error);
      // Return default preferences on error
      return { ...DEFAULT_PREFERENCES, UserId: targetUserId } as IUserPreferences;
    }
  }

  /**
   * Save user preferences
   */
  public async savePreferences(prefs: IUserPreferences): Promise<void> {
    try {
      // Validate user ID
      ValidationUtils.validateUserId(prefs.UserId);

      const serializedPrefs = this.serializePreferences(prefs);

      // Check if preferences exist with secure filter
      const filter = ValidationUtils.buildFilter('UserId', 'eq', prefs.UserId);
      const existing = await this.sp.web.lists
        .getByTitle('JML_UserPreferences')
        .items
        .filter(filter)
        .top(1)();

      if (existing.length > 0) {
        // Update existing preferences
        await this.sp.web.lists
          .getByTitle('JML_UserPreferences')
          .items
          .getById(existing[0].Id)
          .update(serializedPrefs);

        prefs.Id = existing[0].Id;
      } else {
        // Create new preferences
        const result = await this.sp.web.lists
          .getByTitle('JML_UserPreferences')
          .items
          .add(serializedPrefs);

        prefs.Id = result.data.Id;
      }

      // Update cache
      this.cachePreferences(prefs);

      // Log activity
      await this.logActivity(ActivityType.PreferencesUpdated, {
        userId: prefs.UserId,
        changes: Object.keys(prefs)
      });
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to save user preferences:', error);
      throw error;
    }
  }

  /**
   * Reset preferences to defaults
   */
  public async resetToDefaults(): Promise<void> {
    const defaultPrefs: IUserPreferences = {
      ...DEFAULT_PREFERENCES,
      UserId: this.currentUserId,
      UserEmail: this.currentUserEmail,
      DisplayName: ''
    } as IUserPreferences;

    await this.savePreferences(defaultPrefs);
    this.clearCache();
  }

  /**
   * Update dashboard layout
   */
  public async updateDashboardLayout(layout: IDashboardLayout): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.DashboardLayout = layout;
    await this.savePreferences(prefs);
  }

  /**
   * Add favorite process
   */
  public async addFavoriteProcess(processId: number): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.FavoriteProcesses = prefs.FavoriteProcesses || [];

    if (!prefs.FavoriteProcesses.includes(processId)) {
      prefs.FavoriteProcesses.push(processId);
      await this.savePreferences(prefs);
    }
  }

  /**
   * Remove favorite process
   */
  public async removeFavoriteProcess(processId: number): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.FavoriteProcesses = prefs.FavoriteProcesses || [];
    prefs.FavoriteProcesses = prefs.FavoriteProcesses.filter(id => id !== processId);
    await this.savePreferences(prefs);
  }

  /**
   * Add favorite template
   */
  public async addFavoriteTemplate(templateId: number): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.FavoriteTemplates = prefs.FavoriteTemplates || [];

    if (!prefs.FavoriteTemplates.includes(templateId)) {
      prefs.FavoriteTemplates.push(templateId);
      await this.savePreferences(prefs);
    }
  }

  /**
   * Remove favorite template
   */
  public async removeFavoriteTemplate(templateId: number): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.FavoriteTemplates = prefs.FavoriteTemplates || [];
    prefs.FavoriteTemplates = prefs.FavoriteTemplates.filter(id => id !== templateId);
    await this.savePreferences(prefs);
  }

  /**
   * Save filter
   */
  public async saveFilter(filter: ISavedFilter): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.SavedFilters = prefs.SavedFilters || [];

    // Remove existing filter with same ID
    prefs.SavedFilters = prefs.SavedFilters.filter(f => f.id !== filter.id);

    // Add new/updated filter
    prefs.SavedFilters.push(filter);

    await this.savePreferences(prefs);
    await this.logActivity(ActivityType.FilterSaved, { filterName: filter.name });
  }

  /**
   * Delete filter
   */
  public async deleteFilter(filterId: string): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.SavedFilters = prefs.SavedFilters || [];
    prefs.SavedFilters = prefs.SavedFilters.filter(f => f.id !== filterId);
    await this.savePreferences(prefs);
  }

  /**
   * Save view
   */
  public async saveView(view: ISavedView): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.SavedViews = prefs.SavedViews || [];

    // Remove existing view with same ID
    prefs.SavedViews = prefs.SavedViews.filter(v => v.id !== view.id);

    // Add new/updated view
    prefs.SavedViews.push(view);

    await this.savePreferences(prefs);
    await this.logActivity(ActivityType.ViewSaved, { viewName: view.name });
  }

  /**
   * Delete view
   */
  public async deleteView(viewId: string): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.SavedViews = prefs.SavedViews || [];
    prefs.SavedViews = prefs.SavedViews.filter(v => v.id !== viewId);
    await this.savePreferences(prefs);
  }

  /**
   * Set default view
   */
  public async setDefaultView(viewName: string): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.DefaultView = viewName;
    await this.savePreferences(prefs);
  }

  /**
   * Update notification settings
   */
  public async updateNotificationSettings(settings: INotificationSettings): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.NotificationSettings = settings;
    await this.savePreferences(prefs);
  }

  /**
   * Update theme preference
   */
  public async updateThemePreference(theme: Partial<IUserPreferences['ThemePreference']>): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.ThemePreference = {
      ...prefs.ThemePreference,
      ...theme
    };
    await this.savePreferences(prefs);
  }

  /**
   * Set custom theme
   */
  public async setCustomTheme(theme: ICustomTheme): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.CustomTheme = theme;
    await this.savePreferences(prefs);
  }

  /**
   * Update language preference
   */
  public async updateLanguage(language: string): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.Language = language;
    await this.savePreferences(prefs);
  }

  /**
   * Update timezone preference
   */
  public async updateTimeZone(timeZone: string): Promise<void> {
    const prefs = await this.getPreferences();
    prefs.TimeZone = timeZone;
    await this.savePreferences(prefs);
  }

  /**
   * Get department branding
   */
  public async getDepartmentBranding(department: string): Promise<any> {
    try {
      // Validate and sanitize department input
      if (!department || typeof department !== 'string' || department.trim().length === 0) {
        throw new Error('Invalid department');
      }

      // Build secure filter
      const filter = `${ValidationUtils.buildFilter('Department', 'eq', department)} and IsEnabled eq true`;

      const items = await this.sp.web.lists
        .getByTitle('JML_DepartmentBranding')
        .items
        .filter(filter)
        .top(1)();

      return items.length > 0 ? items[0] : null;
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to get department branding:', error);
      return null;
    }
  }

  /**
   * Get favorite items details
   */
  public async getFavorites(): Promise<{
    processes: IFavoriteItem[];
    templates: IFavoriteItem[];
  }> {
    const prefs = await this.getPreferences();
    const favorites = {
      processes: [] as IFavoriteItem[],
      templates: [] as IFavoriteItem[]
    };

    try {
      // Get favorite processes
      if (prefs.FavoriteProcesses && prefs.FavoriteProcesses.length > 0) {
        const processes = await this.sp.web.lists
          .getByTitle('JML_Processes')
          .items
          .filter(`Id in (${prefs.FavoriteProcesses.join(',')})`)
          .select('Id', 'Title', 'ProcessType', 'Created')();

        favorites.processes = processes.map(p => ({
          id: p.Id,
          type: 'process' as const,
          title: p.Title,
          url: `/process/${p.Id}`,
          icon: this.getProcessIcon(p.ProcessType),
          addedDate: new Date(p.Created),
          accessCount: 0
        }));
      }

      // Get favorite templates
      if (prefs.FavoriteTemplates && prefs.FavoriteTemplates.length > 0) {
        const templates = await this.sp.web.lists
          .getByTitle('JML_ProcessChecklistTemplates')
          .items
          .filter(`Id in (${prefs.FavoriteTemplates.join(',')})`)
          .select('Id', 'Title', 'ProcessType', 'Created')();

        favorites.templates = templates.map(t => ({
          id: t.Id,
          type: 'template' as const,
          title: t.Title,
          url: `/template/${t.Id}`,
          icon: this.getProcessIcon(t.ProcessType),
          addedDate: new Date(t.Created),
          accessCount: 0
        }));
      }
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to get favorites:', error);
    }

    return favorites;
  }

  /**
   * Export preferences
   */
  public async exportPreferences(): Promise<string> {
    const prefs = await this.getPreferences();
    return JSON.stringify(prefs, null, 2);
  }

  /**
   * Import preferences
   */
  public async importPreferences(jsonData: string): Promise<void> {
    try {
      // Validate JSON before parsing (prevent DoS and validate format)
      const parsed = ValidationUtils.validateJSON(jsonData, 1048576); // 1MB max

      const prefs = parsed as IUserPreferences;

      // Ensure current user ID (security: prevent importing other users' prefs)
      prefs.UserId = this.currentUserId;
      prefs.UserEmail = this.currentUserEmail;

      await this.savePreferences(prefs);
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to import preferences:', error);
      throw new Error('Invalid preferences data');
    }
  }

  /**
   * Get user activity history
   */
  public async getUserActivity(days: number = 30): Promise<any[]> {
    try {
      // Validate days parameter
      const validDays = ValidationUtils.validateInteger(days, 'days', 1, 365);

      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - validDays);

      // Build secure filter
      const userFilter = ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId);
      const dateFilter = ValidationUtils.buildFilter('Created', 'ge', cutoffDate);
      const filter = `${userFilter} and ${dateFilter}`;

      const activities = await this.sp.web.lists
        .getByTitle('JML_UserActivity')
        .items
        .filter(filter)
        .orderBy('Created', false)
        .top(100)();

      return activities;
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to get user activity:', error);
      return [];
    }
  }

  /**
   * Create default preferences for new user
   */
  private async createDefaultPreferences(userId: string): Promise<IUserPreferences> {
    try {
      // Get user info
      const user = await this.sp.web.siteUsers.getById(parseInt(userId))();

      const defaultPrefs: IUserPreferences = {
        ...DEFAULT_PREFERENCES,
        UserId: userId,
        UserEmail: user.Email,
        DisplayName: user.Title
      } as IUserPreferences;

      // Save to SharePoint
      await this.savePreferences(defaultPrefs);

      return defaultPrefs;
    } catch (error) {
      logger.error('UserPreferencesService', 'Failed to create default preferences:', error);
      return { ...DEFAULT_PREFERENCES, UserId: userId } as IUserPreferences;
    }
  }

  /**
   * Parse preferences from SharePoint item
   */
  private parsePreferences(item: any): IUserPreferences {
    return {
      Id: item.Id,
      Title: item.Title,
      UserId: item.UserId,
      UserEmail: item.UserEmail,
      DisplayName: item.DisplayName,
      DashboardLayout: item.DashboardLayout ? JSON.parse(item.DashboardLayout) : DEFAULT_PREFERENCES.DashboardLayout,
      FavoriteProcesses: item.FavoriteProcesses ? JSON.parse(item.FavoriteProcesses) : [],
      FavoriteTemplates: item.FavoriteTemplates ? JSON.parse(item.FavoriteTemplates) : [],
      FavoriteViews: item.FavoriteViews ? JSON.parse(item.FavoriteViews) : [],
      ThemePreference: item.ThemePreference ? JSON.parse(item.ThemePreference) : DEFAULT_PREFERENCES.ThemePreference,
      CustomTheme: item.CustomTheme ? JSON.parse(item.CustomTheme) : undefined,
      SavedFilters: item.SavedFilters ? JSON.parse(item.SavedFilters) : [],
      SavedViews: item.SavedViews ? JSON.parse(item.SavedViews) : [],
      DefaultView: item.DefaultView,
      NotificationSettings: item.NotificationSettings ? JSON.parse(item.NotificationSettings) : DEFAULT_PREFERENCES.NotificationSettings,
      Language: item.Language || DEFAULT_PREFERENCES.Language,
      TimeZone: item.TimeZone || DEFAULT_PREFERENCES.TimeZone,
      DateFormat: item.DateFormat || DEFAULT_PREFERENCES.DateFormat,
      TimeFormat: item.TimeFormat || DEFAULT_PREFERENCES.TimeFormat,
      NumberFormat: item.NumberFormat,
      DensityMode: item.DensityMode || DEFAULT_PREFERENCES.DensityMode,
      DefaultItemsPerPage: item.DefaultItemsPerPage || DEFAULT_PREFERENCES.DefaultItemsPerPage,
      ShowWelcomeMessage: item.ShowWelcomeMessage ?? DEFAULT_PREFERENCES.ShowWelcomeMessage,
      ShowTipsAndTricks: item.ShowTipsAndTricks ?? DEFAULT_PREFERENCES.ShowTipsAndTricks,
      KeyboardShortcutsEnabled: item.KeyboardShortcutsEnabled ?? DEFAULT_PREFERENCES.KeyboardShortcutsEnabled,
      AnimationsEnabled: item.AnimationsEnabled ?? DEFAULT_PREFERENCES.AnimationsEnabled,
      AccessibilityMode: item.AccessibilityMode ?? DEFAULT_PREFERENCES.AccessibilityMode,
      HighContrastMode: item.HighContrastMode ?? DEFAULT_PREFERENCES.HighContrastMode,
      Preferences: item.Preferences,
      LastModified: item.Modified ? new Date(item.Modified) : undefined,
      Created: item.Created ? new Date(item.Created) : undefined,
      Version: item.Version
    };
  }

  /**
   * Serialize preferences for SharePoint
   */
  private serializePreferences(prefs: IUserPreferences): any {
    return {
      Title: `Preferences - ${prefs.UserEmail}`,
      UserId: prefs.UserId,
      UserEmail: prefs.UserEmail,
      DisplayName: prefs.DisplayName,
      DashboardLayout: prefs.DashboardLayout ? JSON.stringify(prefs.DashboardLayout) : null,
      FavoriteProcesses: prefs.FavoriteProcesses ? JSON.stringify(prefs.FavoriteProcesses) : null,
      FavoriteTemplates: prefs.FavoriteTemplates ? JSON.stringify(prefs.FavoriteTemplates) : null,
      FavoriteViews: prefs.FavoriteViews ? JSON.stringify(prefs.FavoriteViews) : null,
      ThemePreference: prefs.ThemePreference ? JSON.stringify(prefs.ThemePreference) : null,
      CustomTheme: prefs.CustomTheme ? JSON.stringify(prefs.CustomTheme) : null,
      SavedFilters: prefs.SavedFilters ? JSON.stringify(prefs.SavedFilters) : null,
      SavedViews: prefs.SavedViews ? JSON.stringify(prefs.SavedViews) : null,
      DefaultView: prefs.DefaultView,
      NotificationSettings: prefs.NotificationSettings ? JSON.stringify(prefs.NotificationSettings) : null,
      Language: prefs.Language,
      TimeZone: prefs.TimeZone,
      DateFormat: prefs.DateFormat,
      TimeFormat: prefs.TimeFormat,
      NumberFormat: prefs.NumberFormat,
      DensityMode: prefs.DensityMode,
      DefaultItemsPerPage: prefs.DefaultItemsPerPage,
      ShowWelcomeMessage: prefs.ShowWelcomeMessage,
      ShowTipsAndTricks: prefs.ShowTipsAndTricks,
      KeyboardShortcutsEnabled: prefs.KeyboardShortcutsEnabled,
      AnimationsEnabled: prefs.AnimationsEnabled,
      AccessibilityMode: prefs.AccessibilityMode,
      HighContrastMode: prefs.HighContrastMode,
      Preferences: prefs.Preferences
    };
  }

  /**
   * Cache preferences
   */
  private cachePreferences(prefs: IUserPreferences): void {
    this.cachedPreferences = prefs;
    this.lastCacheTime = Date.now();
  }

  /**
   * Clear cache
   */
  private clearCache(): void {
    this.cachedPreferences = null;
    this.lastCacheTime = 0;
  }

  /**
   * Log user activity
   */
  private async logActivity(activityType: ActivityType, metadata?: any): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('JML_UserActivity')
        .items
        .add({
          Title: `${activityType} - ${new Date().toISOString()}`,
          UserId: this.currentUserId,
          UserEmail: this.currentUserEmail,
          ActivityType: activityType,
          Metadata: metadata ? JSON.stringify(metadata) : undefined
        });
    } catch (error) {
      // Don't fail the operation if logging fails
      logger.error('UserPreferencesService', 'Failed to log user activity:', error);
    }
  }

  /**
   * Get process icon
   */
  private getProcessIcon(processType: string): string {
    const icons: Record<string, string> = {
      'Onboarding': '👋',
      'Offboarding': '👋',
      'Internal Transfer': '🔄',
      'Promotion': '⬆️',
      'Leave of Absence': '🏖️'
    };
    return icons[processType] || '📋';
  }

  // Theme Management Methods (Placeholder - will be implemented with ThemeService)
  public async getAllThemes(): Promise<any[]> {
    // TODO: Integrate with ThemeService
    return [];
  }

  public async getActiveTheme(): Promise<any | null> {
    // TODO: Integrate with ThemeService
    return null;
  }

  public async saveTheme(theme: any): Promise<void> {
    // TODO: Integrate with ThemeService
    logger.info('UserPreferencesService', 'saveTheme called (placeholder)');
  }

  public async deleteTheme(themeId: string): Promise<void> {
    // TODO: Integrate with ThemeService
    logger.info('UserPreferencesService', 'deleteTheme called (placeholder)');
  }

  public async duplicateTheme(themeId: string, newName: string): Promise<any> {
    // TODO: Integrate with ThemeService
    logger.info('UserPreferencesService', `duplicateTheme called for ${themeId} -> ${newName} (placeholder)`);
    return null;
  }

  public async setActiveTheme(themeId: string): Promise<void> {
    // TODO: Integrate with ThemeService
    logger.info('UserPreferencesService', 'setActiveTheme called (placeholder)');
  }

  public async importTheme(themeJson: string): Promise<any> {
    // TODO: Integrate with ThemeService
    logger.info('UserPreferencesService', 'importTheme called (placeholder)');
    return null;
  }
}
