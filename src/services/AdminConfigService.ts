// AdminConfigService.ts
// Service layer for Admin Panel configuration — CRUD against PM_NamingRules,
// PM_SLAConfigs, PM_DataLifecyclePolicies, PM_EmailTemplates, plus
// key-value settings via PM_Configuration.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { AdminConfigLists, PolicyLists } from '../constants/SharePointListNames';
import { logger } from './LoggingService';
import {
  INamingRule,
  ISLAConfig,
  IDataLifecyclePolicy,
  IEmailTemplate,
  IPolicyCategory,
  IPolicySubCategory,
  IGeneralSettings,
  IPolicyMetadataProfile,
  AdminConfigKeys
} from '../models/IAdminConfig';
import { IPolicyTemplate } from '../models/IPolicy';

// ============================================================================
// Service
// ============================================================================

export class AdminConfigService {
  private sp: SPFI;

  private readonly NAMING_RULES_LIST = AdminConfigLists.NAMING_RULES;
  private readonly SLA_CONFIGS_LIST = AdminConfigLists.SLA_CONFIGS;
  private readonly LIFECYCLE_LIST = AdminConfigLists.DATA_LIFECYCLE_POLICIES;
  private readonly EMAIL_TEMPLATES_LIST = AdminConfigLists.EMAIL_TEMPLATES;
  private readonly CATEGORIES_LIST = PolicyLists.POLICY_CATEGORIES;
  private readonly SUB_CATEGORIES_LIST = PolicyLists.POLICY_SUB_CATEGORIES;
  private readonly TEMPLATES_LIST = PolicyLists.POLICY_TEMPLATES;
  private readonly METADATA_PROFILES_LIST = PolicyLists.POLICY_METADATA_PROFILES;
  private readonly CONFIG_LIST = 'PM_Configuration';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ──────────── Naming Rules CRUD ────────────

  public async getNamingRules(): Promise<INamingRule[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.NAMING_RULES_LIST)
        .items.select('Id', 'Title', 'Pattern', 'Segments', 'AppliesTo', 'IsActive', 'Example')
        .orderBy('Title')
        .top(100)();

      logger.info('AdminConfigService', `Loaded ${items.length} naming rules`);
      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        Pattern: item.Pattern || '',
        Segments: this.parseJson(item.Segments, []),
        AppliesTo: item.AppliesTo || 'All Policies',
        IsActive: item.IsActive ?? true,
        Example: item.Example || ''
      }));
    } catch (error) {
      logger.error('AdminConfigService', 'getNamingRules failed:', error);
      return [];
    }
  }

  public async createNamingRule(rule: INamingRule): Promise<INamingRule> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.NAMING_RULES_LIST)
        .items.add({
          Title: rule.Title,
          Pattern: rule.Pattern,
          Segments: JSON.stringify(rule.Segments),
          AppliesTo: rule.AppliesTo,
          IsActive: rule.IsActive,
          Example: rule.Example
        });

      const newId = result.data?.Id ?? 0;
      logger.info('AdminConfigService', `Created naming rule id=${newId}`);
      return { ...rule, Id: newId };
    } catch (error) {
      logger.error('AdminConfigService', 'createNamingRule failed:', error);
      throw error;
    }
  }

  public async updateNamingRule(id: number, rule: INamingRule): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.NAMING_RULES_LIST)
        .items.getById(id)
        .update({
          Title: rule.Title,
          Pattern: rule.Pattern,
          Segments: JSON.stringify(rule.Segments),
          AppliesTo: rule.AppliesTo,
          IsActive: rule.IsActive,
          Example: rule.Example
        });

      logger.info('AdminConfigService', `Updated naming rule id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateNamingRule id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteNamingRule(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.NAMING_RULES_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted naming rule id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteNamingRule id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── SLA Configs CRUD ────────────

  public async getSLAConfigs(): Promise<ISLAConfig[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.SLA_CONFIGS_LIST)
        .items.select('Id', 'Title', 'ProcessType', 'TargetDays', 'WarningThresholdDays', 'IsActive', 'Description')
        .orderBy('Title')
        .top(100)();

      logger.info('AdminConfigService', `Loaded ${items.length} SLA configs`);
      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        ProcessType: item.ProcessType || '',
        TargetDays: item.TargetDays || 7,
        WarningThresholdDays: item.WarningThresholdDays || 2,
        IsActive: item.IsActive ?? true,
        Description: item.Description || ''
      }));
    } catch (error) {
      logger.error('AdminConfigService', 'getSLAConfigs failed:', error);
      return [];
    }
  }

  public async createSLAConfig(config: ISLAConfig): Promise<ISLAConfig> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.SLA_CONFIGS_LIST)
        .items.add({
          Title: config.Title,
          ProcessType: config.ProcessType,
          TargetDays: config.TargetDays,
          WarningThresholdDays: config.WarningThresholdDays,
          IsActive: config.IsActive,
          Description: config.Description
        });

      const newId = result.data?.Id ?? 0;
      logger.info('AdminConfigService', `Created SLA config id=${newId}`);
      return { ...config, Id: newId };
    } catch (error) {
      logger.error('AdminConfigService', 'createSLAConfig failed:', error);
      throw error;
    }
  }

  public async updateSLAConfig(id: number, config: ISLAConfig): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.SLA_CONFIGS_LIST)
        .items.getById(id)
        .update({
          Title: config.Title,
          ProcessType: config.ProcessType,
          TargetDays: config.TargetDays,
          WarningThresholdDays: config.WarningThresholdDays,
          IsActive: config.IsActive,
          Description: config.Description
        });

      logger.info('AdminConfigService', `Updated SLA config id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateSLAConfig id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteSLAConfig(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.SLA_CONFIGS_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted SLA config id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteSLAConfig id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Data Lifecycle CRUD ────────────

  public async getLifecyclePolicies(): Promise<IDataLifecyclePolicy[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LIFECYCLE_LIST)
        .items.select('Id', 'Title', 'EntityType', 'RetentionPeriodDays', 'AutoDeleteEnabled', 'ArchiveBeforeDelete', 'IsActive', 'Description')
        .orderBy('Title')
        .top(100)();

      logger.info('AdminConfigService', `Loaded ${items.length} lifecycle policies`);
      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        EntityType: item.EntityType || '',
        RetentionPeriodDays: item.RetentionPeriodDays || 365,
        AutoDeleteEnabled: item.AutoDeleteEnabled ?? false,
        ArchiveBeforeDelete: item.ArchiveBeforeDelete ?? true,
        IsActive: item.IsActive ?? true,
        Description: item.Description || ''
      }));
    } catch (error) {
      logger.error('AdminConfigService', 'getLifecyclePolicies failed:', error);
      return [];
    }
  }

  public async createLifecyclePolicy(policy: IDataLifecyclePolicy): Promise<IDataLifecyclePolicy> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.LIFECYCLE_LIST)
        .items.add({
          Title: policy.Title,
          EntityType: policy.EntityType,
          RetentionPeriodDays: policy.RetentionPeriodDays,
          AutoDeleteEnabled: policy.AutoDeleteEnabled,
          ArchiveBeforeDelete: policy.ArchiveBeforeDelete,
          IsActive: policy.IsActive,
          Description: policy.Description
        });

      const newId = result.data?.Id ?? 0;
      logger.info('AdminConfigService', `Created lifecycle policy id=${newId}`);
      return { ...policy, Id: newId };
    } catch (error) {
      logger.error('AdminConfigService', 'createLifecyclePolicy failed:', error);
      throw error;
    }
  }

  public async updateLifecyclePolicy(id: number, policy: IDataLifecyclePolicy): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.LIFECYCLE_LIST)
        .items.getById(id)
        .update({
          Title: policy.Title,
          EntityType: policy.EntityType,
          RetentionPeriodDays: policy.RetentionPeriodDays,
          AutoDeleteEnabled: policy.AutoDeleteEnabled,
          ArchiveBeforeDelete: policy.ArchiveBeforeDelete,
          IsActive: policy.IsActive,
          Description: policy.Description
        });

      logger.info('AdminConfigService', `Updated lifecycle policy id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateLifecyclePolicy id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteLifecyclePolicy(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.LIFECYCLE_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted lifecycle policy id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteLifecyclePolicy id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Email Templates CRUD ────────────

  public async getEmailTemplates(): Promise<IEmailTemplate[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.EMAIL_TEMPLATES_LIST)
        .items.select('Id', 'Title', 'EventTrigger', 'Subject', 'Body', 'Recipients', 'IsActive', 'MergeTags', 'Modified')
        .orderBy('Title')
        .top(100)();

      logger.info('AdminConfigService', `Loaded ${items.length} email templates`);
      return items.map((item: any) => ({
        id: item.Id,
        name: item.Title || '',
        event: item.EventTrigger || '',
        subject: item.Subject || '',
        body: item.Body || '',
        recipients: item.Recipients || '',
        isActive: item.IsActive ?? true,
        mergeTags: this.parseJson(item.MergeTags, []),
        lastModified: item.Modified ? new Date(item.Modified).toISOString().split('T')[0] : ''
      }));
    } catch (error) {
      logger.error('AdminConfigService', 'getEmailTemplates failed:', error);
      return [];
    }
  }

  public async createEmailTemplate(template: IEmailTemplate): Promise<IEmailTemplate> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.EMAIL_TEMPLATES_LIST)
        .items.add({
          Title: template.name,
          EventTrigger: template.event,
          Subject: template.subject,
          Body: template.body,
          Recipients: template.recipients,
          IsActive: template.isActive,
          MergeTags: JSON.stringify(template.mergeTags)
        });

      const newId = result.data?.Id ?? 0;
      logger.info('AdminConfigService', `Created email template id=${newId}`);
      return { ...template, id: newId };
    } catch (error) {
      logger.error('AdminConfigService', 'createEmailTemplate failed:', error);
      throw error;
    }
  }

  public async updateEmailTemplate(id: number, template: IEmailTemplate): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.EMAIL_TEMPLATES_LIST)
        .items.getById(id)
        .update({
          Title: template.name,
          EventTrigger: template.event,
          Subject: template.subject,
          Body: template.body,
          Recipients: template.recipients,
          IsActive: template.isActive,
          MergeTags: JSON.stringify(template.mergeTags)
        });

      logger.info('AdminConfigService', `Updated email template id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateEmailTemplate id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteEmailTemplate(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.EMAIL_TEMPLATES_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted email template id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteEmailTemplate id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Policy Categories CRUD ────────────

  public async getCategories(): Promise<IPolicyCategory[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CATEGORIES_LIST)
        .items.select('Id', 'Title', 'CategoryName', 'IconName', 'Color', 'Description', 'SortOrder', 'IsActive', 'IsDefault')
        .orderBy('SortOrder')
        .top(100)();

      logger.info('AdminConfigService', `Loaded ${items.length} policy categories`);
      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        CategoryName: item.CategoryName || item.Title || '',
        IconName: item.IconName || 'Tag',
        Color: item.Color || '#0d9488',
        Description: item.Description || '',
        SortOrder: item.SortOrder ?? 99,
        IsActive: item.IsActive ?? true,
        IsDefault: item.IsDefault ?? false
      }));
    } catch (error) {
      logger.error('AdminConfigService', 'getCategories failed:', error);
      return [];
    }
  }

  public async createCategory(category: IPolicyCategory): Promise<IPolicyCategory> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.CATEGORIES_LIST)
        .items.add({
          Title: category.CategoryName,
          CategoryName: category.CategoryName,
          IconName: category.IconName,
          Color: category.Color,
          Description: category.Description,
          SortOrder: category.SortOrder,
          IsActive: category.IsActive,
          IsDefault: false
        });

      const newId = result.data?.Id ?? 0;
      logger.info('AdminConfigService', `Created category id=${newId}: ${category.CategoryName}`);
      return { ...category, Id: newId, IsDefault: false };
    } catch (error) {
      logger.error('AdminConfigService', 'createCategory failed:', error);
      throw error;
    }
  }

  public async updateCategory(id: number, category: IPolicyCategory): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.CATEGORIES_LIST)
        .items.getById(id)
        .update({
          Title: category.CategoryName,
          CategoryName: category.CategoryName,
          IconName: category.IconName,
          Color: category.Color,
          Description: category.Description,
          SortOrder: category.SortOrder,
          IsActive: category.IsActive
        });

      logger.info('AdminConfigService', `Updated category id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateCategory id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteCategory(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.CATEGORIES_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted category id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteCategory id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Sub-Categories ────────────

  public async getSubCategories(parentCategoryId?: number): Promise<IPolicySubCategory[]> {
    try {
      let query = this.sp.web.lists
        .getByTitle(this.SUB_CATEGORIES_LIST)
        .items.select('Id', 'Title', 'SubCategoryName', 'ParentCategoryId', 'ParentCategoryName', 'IconName', 'Description', 'SortOrder', 'IsActive')
        .orderBy('SortOrder')
        .top(200);

      if (parentCategoryId) {
        query = query.filter(`ParentCategoryId eq ${parentCategoryId}`);
      }

      const items = await query();
      logger.info('AdminConfigService', `Loaded ${items.length} sub-categories`);
      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        SubCategoryName: item.SubCategoryName || item.Title || '',
        ParentCategoryId: item.ParentCategoryId || 0,
        ParentCategoryName: item.ParentCategoryName || '',
        IconName: item.IconName || 'FolderOpen',
        Description: item.Description || '',
        SortOrder: item.SortOrder ?? 99,
        IsActive: item.IsActive ?? true
      }));
    } catch (error) {
      logger.error('AdminConfigService', 'getSubCategories failed:', error);
      return [];
    }
  }

  public async createSubCategory(subCategory: IPolicySubCategory): Promise<IPolicySubCategory> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.SUB_CATEGORIES_LIST)
        .items.add({
          Title: subCategory.SubCategoryName,
          SubCategoryName: subCategory.SubCategoryName,
          ParentCategoryId: subCategory.ParentCategoryId,
          ParentCategoryName: subCategory.ParentCategoryName,
          IconName: subCategory.IconName,
          Description: subCategory.Description,
          SortOrder: subCategory.SortOrder,
          IsActive: subCategory.IsActive
        });

      const newId = result.data?.Id ?? 0;
      logger.info('AdminConfigService', `Created sub-category id=${newId}: ${subCategory.SubCategoryName}`);
      return { ...subCategory, Id: newId };
    } catch (error) {
      logger.error('AdminConfigService', 'createSubCategory failed:', error);
      throw error;
    }
  }

  public async updateSubCategory(id: number, subCategory: IPolicySubCategory): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.SUB_CATEGORIES_LIST)
        .items.getById(id)
        .update({
          Title: subCategory.SubCategoryName,
          SubCategoryName: subCategory.SubCategoryName,
          ParentCategoryId: subCategory.ParentCategoryId,
          ParentCategoryName: subCategory.ParentCategoryName,
          IconName: subCategory.IconName,
          Description: subCategory.Description,
          SortOrder: subCategory.SortOrder,
          IsActive: subCategory.IsActive
        });

      logger.info('AdminConfigService', `Updated sub-category id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateSubCategory id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteSubCategory(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.SUB_CATEGORIES_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted sub-category id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteSubCategory id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Policy Templates (existing list) ────────────

  public async getTemplates(): Promise<IPolicyTemplate[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.select('Id', 'Title', 'TemplateName', 'TemplateCategory', 'TemplateDescription', 'HTMLTemplate', 'IsActive')
        .orderBy('Title')
        .top(200)();

      logger.info('AdminConfigService', `Loaded ${items.length} policy templates`);
      return items as IPolicyTemplate[];
    } catch (error) {
      logger.error('AdminConfigService', 'getTemplates failed:', error);
      return [];
    }
  }

  public async createTemplate(data: Record<string, unknown>): Promise<any> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.add(data);

      logger.info('AdminConfigService', `Created template id=${result.data?.Id}`);
      return result;
    } catch (error) {
      logger.error('AdminConfigService', 'createTemplate failed:', error);
      throw error;
    }
  }

  public async updateTemplate(id: number, data: Record<string, unknown>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.getById(id)
        .update(data);

      logger.info('AdminConfigService', `Updated template id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateTemplate id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteTemplate(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted template id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteTemplate id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Metadata Profiles (existing list) ────────────

  public async getMetadataProfiles(): Promise<IPolicyMetadataProfile[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.METADATA_PROFILES_LIST)
        .items.select(
          'Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk',
          'ReadTimeframe', 'RequiresAcknowledgement', 'RequiresQuiz',
          'TargetDepartments', 'TargetRoles', 'IsActive', 'Description'
        )
        .orderBy('Title')
        .top(200)();

      logger.info('AdminConfigService', `Loaded ${items.length} metadata profiles`);
      return items as IPolicyMetadataProfile[];
    } catch (error) {
      logger.error('AdminConfigService', 'getMetadataProfiles failed:', error);
      return [];
    }
  }

  public async createMetadataProfile(data: Record<string, unknown>): Promise<any> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.METADATA_PROFILES_LIST)
        .items.add(data);

      logger.info('AdminConfigService', `Created metadata profile id=${result.data?.Id}`);
      return result;
    } catch (error) {
      logger.error('AdminConfigService', 'createMetadataProfile failed:', error);
      throw error;
    }
  }

  public async updateMetadataProfile(id: number, data: Record<string, unknown>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.METADATA_PROFILES_LIST)
        .items.getById(id)
        .update(data);

      logger.info('AdminConfigService', `Updated metadata profile id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `updateMetadataProfile id=${id} failed:`, error);
      throw error;
    }
  }

  public async deleteMetadataProfile(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.METADATA_PROFILES_LIST)
        .items.getById(id)
        .delete();

      logger.info('AdminConfigService', `Deleted metadata profile id=${id}`);
    } catch (error) {
      logger.error('AdminConfigService', `deleteMetadataProfile id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── General Settings (PM_Configuration key-value) ────────────

  public async getGeneralSettings(): Promise<Partial<IGeneralSettings>> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items.filter("substringof('Admin.General', ConfigKey)")
        .select('Id', 'ConfigKey', 'ConfigValue')
        .top(20)();

      const settings: Record<string, string> = {};
      items.forEach((item: any) => {
        settings[item.ConfigKey] = item.ConfigValue;
      });

      logger.info('AdminConfigService', `Loaded ${items.length} general settings from PM_Configuration`);

      return {
        showFeaturedPolicy: settings[AdminConfigKeys.GENERAL_SHOW_FEATURED] === 'true',
        showRecentlyViewed: settings[AdminConfigKeys.GENERAL_SHOW_RECENTLY_VIEWED] !== 'false',
        showQuickStats: settings[AdminConfigKeys.GENERAL_SHOW_QUICK_STATS] !== 'false',
        defaultViewMode: (settings[AdminConfigKeys.GENERAL_DEFAULT_VIEW] as 'table' | 'card') || undefined,
        policiesPerPage: settings[AdminConfigKeys.GENERAL_POLICIES_PER_PAGE] ? Number(settings[AdminConfigKeys.GENERAL_POLICIES_PER_PAGE]) : undefined,
        enableSocialFeatures: settings[AdminConfigKeys.GENERAL_SOCIAL_FEATURES] !== 'false',
        enablePolicyRatings: settings[AdminConfigKeys.GENERAL_POLICY_RATINGS] !== 'false',
        enablePolicyComments: settings[AdminConfigKeys.GENERAL_POLICY_COMMENTS] !== 'false',
        maintenanceMode: settings[AdminConfigKeys.GENERAL_MAINTENANCE_MODE] === 'true',
        maintenanceMessage: settings[AdminConfigKeys.GENERAL_MAINTENANCE_MESSAGE] || undefined
      };
    } catch (error) {
      logger.error('AdminConfigService', 'getGeneralSettings failed:', error);
      return {};
    }
  }

  public async saveGeneralSettings(settings: IGeneralSettings): Promise<void> {
    const pairs: Array<{ key: string; value: string }> = [
      { key: AdminConfigKeys.GENERAL_SHOW_FEATURED, value: String(settings.showFeaturedPolicy) },
      { key: AdminConfigKeys.GENERAL_SHOW_RECENTLY_VIEWED, value: String(settings.showRecentlyViewed) },
      { key: AdminConfigKeys.GENERAL_SHOW_QUICK_STATS, value: String(settings.showQuickStats) },
      { key: AdminConfigKeys.GENERAL_DEFAULT_VIEW, value: settings.defaultViewMode },
      { key: AdminConfigKeys.GENERAL_POLICIES_PER_PAGE, value: String(settings.policiesPerPage) },
      { key: AdminConfigKeys.GENERAL_SOCIAL_FEATURES, value: String(settings.enableSocialFeatures) },
      { key: AdminConfigKeys.GENERAL_POLICY_RATINGS, value: String(settings.enablePolicyRatings) },
      { key: AdminConfigKeys.GENERAL_POLICY_COMMENTS, value: String(settings.enablePolicyComments) },
      { key: AdminConfigKeys.GENERAL_MAINTENANCE_MODE, value: String(settings.maintenanceMode) },
      { key: AdminConfigKeys.GENERAL_MAINTENANCE_MESSAGE, value: settings.maintenanceMessage }
    ];

    for (const pair of pairs) {
      await this.upsertConfigValue(pair.key, pair.value, 'General');
    }

    logger.info('AdminConfigService', 'Saved general settings to PM_Configuration');
  }

  // ──────────── Category-based config (Approval, Compliance, Notifications, Security) ────────────

  public async getConfigByCategory(category: string): Promise<Record<string, string>> {
    try {
      const prefix = `Admin.${category}`;
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items.filter(`substringof('${prefix}', ConfigKey)`)
        .select('Id', 'ConfigKey', 'ConfigValue')
        .top(20)();

      const result: Record<string, string> = {};
      items.forEach((item: any) => {
        result[item.ConfigKey] = item.ConfigValue;
      });

      logger.info('AdminConfigService', `Loaded ${items.length} config values for category '${category}'`);
      return result;
    } catch (error) {
      logger.error('AdminConfigService', `getConfigByCategory '${category}' failed:`, error);
      return {};
    }
  }

  public async saveConfigByCategory(category: string, values: Record<string, string>): Promise<void> {
    for (const [key, value] of Object.entries(values)) {
      await this.upsertConfigValue(key, value, category);
    }
    logger.info('AdminConfigService', `Saved ${Object.keys(values).length} config values for category '${category}'`);
  }

  // ──────────── Private helpers ────────────

  private async upsertConfigValue(key: string, value: string, category: string): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items.filter(`ConfigKey eq '${key}'`)
        .select('Id')
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.CONFIG_LIST)
          .items.getById(items[0].Id)
          .update({ ConfigValue: value });
      } else {
        await this.sp.web.lists
          .getByTitle(this.CONFIG_LIST)
          .items.add({
            Title: key,
            ConfigKey: key,
            ConfigValue: value,
            Category: category,
            IsActive: true,
            IsSystemConfig: false
          });
      }
    } catch (error) {
      logger.error('AdminConfigService', `upsertConfigValue '${key}' failed:`, error);
    }
  }

  private parseJson<T>(value: string | null | undefined, fallback: T): T {
    if (!value) return fallback;
    try {
      return JSON.parse(value);
    } catch {
      return fallback;
    }
  }
}
