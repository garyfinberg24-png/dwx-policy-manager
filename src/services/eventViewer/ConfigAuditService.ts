/**
 * ConfigAuditService — Reads all PM_Configuration values and returns them
 * in a structured format for the Config Audit view. Groups by category,
 * flags missing required keys, and shows overridden defaults.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================================================
// TYPES
// ============================================================================

export interface IConfigEntry {
  id: number;
  key: string;
  value: string;
  category: string;
  isActive: boolean;
  isSystem: boolean;
  /** Whether this is a required key that must have a value */
  isRequired: boolean;
  /** Default value (if known) */
  defaultValue: string;
  /** Whether the current value differs from the default */
  isOverridden: boolean;
}

export interface IConfigAuditSummary {
  entries: IConfigEntry[];
  totalKeys: number;
  activeKeys: number;
  requiredMissing: number;
  overriddenCount: number;
  categories: string[];
  durationMs: number;
}

// ============================================================================
// KNOWN DEFAULTS — what the application uses if no config value is set
// ============================================================================

const KNOWN_DEFAULTS: Record<string, { defaultValue: string; required: boolean }> = {
  // General
  'Admin.General.ShowFeaturedPolicy': { defaultValue: 'true', required: false },
  'Admin.General.ShowRecentlyViewed': { defaultValue: 'true', required: false },
  'Admin.General.ShowQuickStats': { defaultValue: 'true', required: false },
  'Admin.General.DefaultViewMode': { defaultValue: 'grid', required: true },
  'Admin.General.PoliciesPerPage': { defaultValue: '20', required: false },
  'Admin.General.EnableSocialFeatures': { defaultValue: 'false', required: false },
  'Admin.General.EnablePolicyRatings': { defaultValue: 'false', required: false },
  'Admin.General.EnablePolicyComments': { defaultValue: 'false', required: false },
  'Admin.General.MaintenanceMode': { defaultValue: 'false', required: false },

  // Approval
  'Admin.Approval.RequireForNew': { defaultValue: 'true', required: true },
  'Admin.Approval.RequireForUpdate': { defaultValue: 'true', required: false },
  'Admin.Approval.AllowSelfApproval': { defaultValue: 'false', required: false },

  // Compliance
  'Admin.Compliance.RequireAcknowledgement': { defaultValue: 'true', required: true },
  'Admin.Compliance.DefaultDeadlineDays': { defaultValue: '30', required: true },
  'Admin.Compliance.SendReminders': { defaultValue: 'true', required: false },
  'Admin.Compliance.DefaultReviewFrequency': { defaultValue: '12', required: false },
  'Admin.Compliance.SendReviewReminders': { defaultValue: 'true', required: false },

  // Notifications
  'Admin.Notifications.NewPolicies': { defaultValue: 'true', required: false },
  'Admin.Notifications.PolicyUpdates': { defaultValue: 'true', required: false },
  'Admin.Notifications.DailyDigest': { defaultValue: 'false', required: false },

  // AI
  'Integration.AI.Chat.Enabled': { defaultValue: 'false', required: false },
  'Integration.AI.Chat.FunctionUrl': { defaultValue: '', required: false },
  'Integration.AI.Chat.MaxTokens': { defaultValue: '1000', required: false },

  // Event Viewer
  'Admin.EventViewer.Enabled': { defaultValue: 'true', required: true },
  'Admin.EventViewer.AppBufferSize': { defaultValue: '1000', required: false },
  'Admin.EventViewer.ConsoleBufferSize': { defaultValue: '500', required: false },
  'Admin.EventViewer.NetworkBufferSize': { defaultValue: '500', required: false },
  'Admin.EventViewer.AutoPersistThreshold': { defaultValue: 'Error', required: false },
  'Admin.EventViewer.AITriageEnabled': { defaultValue: 'false', required: false },
  'Admin.EventViewer.AIFunctionUrl': { defaultValue: '', required: false },
  'Admin.EventViewer.RetentionDays': { defaultValue: '90', required: false },
  'Admin.EventViewer.HideCDNByDefault': { defaultValue: 'true', required: false },

  // Performance
  'Perf.CacheTTL': { defaultValue: '30', required: false },
  'Perf.RequestDedup': { defaultValue: 'true', required: false },
  'Perf.LeanQueries': { defaultValue: 'false', required: false },
  'Perf.DefaultTopLimit': { defaultValue: '100', required: false },
  'Perf.MaxConcurrent': { defaultValue: '4', required: false },
};

// ============================================================================
// SERVICE
// ============================================================================

export class ConfigAuditService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Load all config values and produce an audit summary.
   */
  public async audit(): Promise<IConfigAuditSummary> {
    const startTime = Date.now();

    const items = await this._sp.web.lists.getByTitle('PM_Configuration')
      .items
      .select('Id', 'ConfigKey', 'ConfigValue', 'Category', 'IsActive', 'IsSystemConfig')
      .top(500)();

    // Build entries from SP data
    const spMap = new Map<string, typeof items[0]>();
    for (const item of items) {
      spMap.set(item.ConfigKey, item);
    }

    const entries: IConfigEntry[] = [];
    const seenKeys = new Set<string>();

    // First: include all items from SP
    for (const item of items) {
      const key = item.ConfigKey || '';
      seenKeys.add(key);
      const known = KNOWN_DEFAULTS[key];

      entries.push({
        id: item.Id,
        key,
        value: item.ConfigValue || '',
        category: this._deriveCategory(key, item.Category),
        isActive: item.IsActive !== false,
        isSystem: item.IsSystemConfig === true,
        isRequired: known?.required || false,
        defaultValue: known?.defaultValue || '',
        isOverridden: known ? (item.ConfigValue || '') !== known.defaultValue : false,
      });
    }

    // Second: include known defaults that are missing from SP
    for (const [key, def] of Object.entries(KNOWN_DEFAULTS)) {
      if (!seenKeys.has(key)) {
        entries.push({
          id: 0,
          key,
          value: '',
          category: this._deriveCategory(key),
          isActive: false,
          isSystem: false,
          isRequired: def.required,
          defaultValue: def.defaultValue,
          isOverridden: false,
        });
      }
    }

    // Sort by category then key
    entries.sort((a, b) => a.category.localeCompare(b.category) || a.key.localeCompare(b.key));

    const categorySet = new Set<string>();
    entries.forEach(e => categorySet.add(e.category));
    const categories = Array.from(categorySet).sort();
    const requiredMissing = entries.filter(e => e.isRequired && !e.value).length;
    const overriddenCount = entries.filter(e => e.isOverridden).length;

    return {
      entries,
      totalKeys: entries.length,
      activeKeys: entries.filter(e => e.isActive).length,
      requiredMissing,
      overriddenCount,
      categories,
      durationMs: Date.now() - startTime,
    };
  }

  /**
   * Derive category from key prefix or SP Category field.
   */
  private _deriveCategory(key: string, spCategory?: string): string {
    if (spCategory) return spCategory;

    if (key.startsWith('Admin.General')) return 'General';
    if (key.startsWith('Admin.Approval')) return 'Approval';
    if (key.startsWith('Admin.Compliance')) return 'Compliance';
    if (key.startsWith('Admin.Notifications')) return 'Notifications';
    if (key.startsWith('Admin.Security')) return 'Security';
    if (key.startsWith('Admin.EventViewer')) return 'Event Viewer';
    if (key.startsWith('Integration.AI')) return 'AI Integration';
    if (key.startsWith('Perf.')) return 'Performance';
    return 'Other';
  }
}
