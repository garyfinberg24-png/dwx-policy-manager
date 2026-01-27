// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

/**
 * User preference keys for JML application
 */
export enum UserPreferenceKey {
  SkipSplashScreen = 'SkipSplashScreen',
  DashboardLayout = 'DashboardLayout',
  Theme = 'Theme',
  DefaultView = 'DefaultView'
}

/**
 * User preference item from SharePoint list
 */
export interface IUserPreference {
  Id?: number;
  Title: string; // User login name
  PreferenceKey: string;
  PreferenceValue: string;
  Modified?: string;
}

/**
 * Service for managing user preferences in SharePoint
 * Uses a SharePoint list called "JML_UserPreferences" to store preferences
 */
export class UserPreferenceService {
  private sp: SPFI;
  private listName: string = 'JML_UserPreferences';
  private currentUserLoginName: string;
  private cache: Map<string, string> = new Map();

  constructor(sp: SPFI, userLoginName: string) {
    this.sp = sp;
    this.currentUserLoginName = userLoginName;
  }

  /**
   * Ensure the preferences list exists with required columns
   */
  public async ensurePreferencesList(): Promise<void> {
    try {
      // Check if list exists
      const lists = await this.sp.web.lists.filter(`Title eq '${this.listName}'`)();

      if (lists.length === 0) {
        console.log(`[UserPreferenceService] Creating list: ${this.listName}`);

        // Create the list
        const listResult = await this.sp.web.lists.add(this.listName, 'Stores user preferences for JML application', 100, false);

        // Add custom fields
        await listResult.list.fields.addText('PreferenceKey', { MaxLength: 255, Required: true });
        await listResult.list.fields.addText('PreferenceValue', { MaxLength: 1000, Required: true });

        console.log(`[UserPreferenceService] List ${this.listName} created successfully`);
      } else {
        // List exists - ensure required columns exist
        const list = this.sp.web.lists.getByTitle(this.listName);
        const fields = await list.fields();
        const fieldNames = fields.map(f => f.InternalName);

        // Add PreferenceKey column if missing
        if (!fieldNames.includes('PreferenceKey')) {
          console.log(`[UserPreferenceService] Adding missing PreferenceKey column`);
          try {
            await list.fields.addText('PreferenceKey', { MaxLength: 255, Required: false });
          } catch (fieldError) {
            console.warn(`[UserPreferenceService] Could not add PreferenceKey column:`, fieldError);
          }
        }

        // Add PreferenceValue column if missing
        if (!fieldNames.includes('PreferenceValue')) {
          console.log(`[UserPreferenceService] Adding missing PreferenceValue column`);
          try {
            await list.fields.addText('PreferenceValue', { MaxLength: 1000, Required: false });
          } catch (fieldError) {
            console.warn(`[UserPreferenceService] Could not add PreferenceValue column:`, fieldError);
          }
        }
      }
    } catch (error) {
      console.error(`[UserPreferenceService] Error ensuring list exists:`, error);
      // Don't throw - fall back to localStorage if list creation fails
    }
  }

  /**
   * Get a preference value for the current user
   */
  public async getPreference(key: UserPreferenceKey): Promise<string | null> {
    // Check cache first
    const cacheKey = `${this.currentUserLoginName}_${key}`;
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey) || null;
    }

    try {
      // Query SharePoint list
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .filter(`Title eq '${this.currentUserLoginName}' and PreferenceKey eq '${key}'`)
        .top(1)();

      if (items.length > 0) {
        const value = items[0].PreferenceValue;
        this.cache.set(cacheKey, value);
        return value;
      }

      return null;
    } catch (error) {
      console.error(`[UserPreferenceService] Error getting preference ${key}:`, error);

      // Fallback to localStorage
      const localValue = localStorage.getItem(cacheKey);
      return localValue;
    }
  }

  /**
   * Set a preference value for the current user
   */
  public async setPreference(key: UserPreferenceKey, value: string): Promise<void> {
    const cacheKey = `${this.currentUserLoginName}_${key}`;

    try {
      // Check if preference already exists
      const existingItems = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .filter(`Title eq '${this.currentUserLoginName}' and PreferenceKey eq '${key}'`)
        .top(1)();

      if (existingItems.length > 0) {
        // Update existing preference
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items
          .getById(existingItems[0].Id)
          .update({
            PreferenceValue: value
          });
      } else {
        // Create new preference
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items
          .add({
            Title: this.currentUserLoginName,
            PreferenceKey: key,
            PreferenceValue: value
          });
      }

      // Update cache
      this.cache.set(cacheKey, value);

      // Also save to localStorage as backup
      localStorage.setItem(cacheKey, value);

      console.log(`[UserPreferenceService] Preference ${key} saved successfully`);
    } catch (error) {
      console.error(`[UserPreferenceService] Error setting preference ${key}:`, error);

      // Fallback to localStorage
      localStorage.setItem(cacheKey, value);
    }
  }

  /**
   * Remove a preference for the current user
   */
  public async removePreference(key: UserPreferenceKey): Promise<void> {
    const cacheKey = `${this.currentUserLoginName}_${key}`;

    try {
      // Find and delete the preference
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .filter(`Title eq '${this.currentUserLoginName}' and PreferenceKey eq '${key}'`)
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items
          .getById(items[0].Id)
          .delete();
      }

      // Remove from cache
      this.cache.delete(cacheKey);

      // Remove from localStorage
      localStorage.removeItem(cacheKey);

      console.log(`[UserPreferenceService] Preference ${key} removed successfully`);
    } catch (error) {
      console.error(`[UserPreferenceService] Error removing preference ${key}:`, error);

      // Fallback to localStorage
      localStorage.removeItem(cacheKey);
    }
  }

  /**
   * Get all preferences for the current user
   */
  public async getAllPreferences(): Promise<Map<string, string>> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .filter(`Title eq '${this.currentUserLoginName}'`)();

      const preferences = new Map<string, string>();
      items.forEach(item => {
        preferences.set(item.PreferenceKey, item.PreferenceValue);

        // Update cache
        const cacheKey = `${this.currentUserLoginName}_${item.PreferenceKey}`;
        this.cache.set(cacheKey, item.PreferenceValue);
      });

      return preferences;
    } catch (error) {
      console.error(`[UserPreferenceService] Error getting all preferences:`, error);
      return new Map();
    }
  }

  /**
   * Clear the in-memory cache
   */
  public clearCache(): void {
    this.cache.clear();
  }

  /**
   * Helper: Check if splash screen should be skipped
   */
  public async shouldSkipSplashScreen(): Promise<boolean> {
    const value = await this.getPreference(UserPreferenceKey.SkipSplashScreen);
    return value === 'true';
  }

  /**
   * Helper: Set splash screen skip preference
   */
  public async setSkipSplashScreen(skip: boolean): Promise<void> {
    await this.setPreference(UserPreferenceKey.SkipSplashScreen, skip.toString());
  }
}
