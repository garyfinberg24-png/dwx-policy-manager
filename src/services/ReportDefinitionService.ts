// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import { IReportDefinition, ReportCategory, IReportLayout } from '../models/IReportBuilder';
import { logger } from './LoggingService';

export class ReportDefinitionService {
  private sp: SPFI;
  private listName: string = 'PM_ReportDefinitions';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get all report definitions
   */
  public async getReportDefinitions(): Promise<IReportDefinition[]> {
    try {
      console.log('[ReportDefinitionService] getReportDefinitions - Loading from list:', this.listName);
      console.log('[ReportDefinitionService] SP object:', this.sp ? 'initialized' : 'NOT initialized');

      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select(
          'ID',
          'ReportName',
          'Description',
          'Category',
          'LayoutConfig',
          'WidgetsJSON',
          'SettingsJSON',
          'GlobalFiltersJSON',
          'IsPublic',
          'Tags',
          'IsTemplate',
          'Status',
          'Created',
          'Modified',
          'Author/Title'
        )
        .expand('Author')
        .orderBy('Modified', false)();

      console.log(`[ReportDefinitionService] Retrieved ${items.length} report definitions from ${this.listName}`);
      return items.map(item => this.mapItemToReportDefinition(item));
    } catch (error: any) {
      console.error(`[ReportDefinitionService] Error loading reports from ${this.listName}:`, error?.message || error);
      logger.error('ReportDefinitionService.getReportDefinitions', error);
      throw error;
    }
  }

  /**
   * Get report definition by ID
   */
  public async getReportById(reportId: number): Promise<IReportDefinition | null> {
    try {
      console.log(`[ReportDefinitionService] getReportById - Loading report ID ${reportId} from ${this.listName}`);

      const item = await this.sp.web.lists.getByTitle(this.listName).items
        .getById(reportId)
        .select(
          'ID',
          'ReportName',
          'Description',
          'Category',
          'LayoutConfig',
          'WidgetsJSON',
          'SettingsJSON',
          'GlobalFiltersJSON',
          'IsPublic',
          'Tags',
          'IsTemplate',
          'Status',
          'Created',
          'Modified',
          'Author/Title'
        )
        .expand('Author')();

      console.log(`[ReportDefinitionService] Retrieved report: ${item.ReportName}`);
      return this.mapItemToReportDefinition(item);
    } catch (error: any) {
      console.error(`[ReportDefinitionService] Error loading report ID ${reportId}:`, error?.message || error);
      logger.error('ReportDefinitionService.getReportById', error);
      return null;
    }
  }

  /**
   * Create a new report definition
   */
  public async createReport(report: Partial<IReportDefinition>): Promise<IReportDefinition> {
    try {
      const itemData = {
        ReportName: report.Title,
        Description: report.Description,
        Category: report.Category,
        LayoutConfig: JSON.stringify(report.layout),
        WidgetsJSON: JSON.stringify(report.widgets || []),
        SettingsJSON: JSON.stringify(report.settings || {}),
        GlobalFiltersJSON: JSON.stringify(report.globalFilters || {}),
        IsPublic: report.isPublic || false,
        Tags: report.tags?.join(';'),
        IsTemplate: false,
        Status: 'Draft'
      };

      const addResult = await this.sp.web.lists.getByTitle(this.listName).items.add(itemData);

      // Fetch the newly created item with all fields
      const newItem = await this.sp.web.lists.getByTitle(this.listName).items
        .getById(addResult.data.ID)
        .select(
          'ID',
          'ReportName',
          'Description',
          'Category',
          'LayoutConfig',
          'WidgetsJSON',
          'SettingsJSON',
          'GlobalFiltersJSON',
          'IsPublic',
          'Status',
          'Created',
          'Modified',
          'Author/Title'
        )
        .expand('Author')();

      logger.info('ReportDefinitionService', `Created report: ${report.Title}`);
      return this.mapItemToReportDefinition(newItem);
    } catch (error) {
      logger.error('ReportDefinitionService.createReport', error);
      throw error;
    }
  }

  /**
   * Update an existing report definition
   */
  public async updateReport(reportId: number, updates: Partial<IReportDefinition>): Promise<void> {
    try {
      const itemData: any = {};

      if (updates.Title) itemData.ReportName = updates.Title;
      if (updates.Description !== undefined) itemData.Description = updates.Description;
      if (updates.Category) itemData.Category = updates.Category;
      if (updates.layout) itemData.LayoutConfig = JSON.stringify(updates.layout);
      if (updates.widgets) itemData.WidgetsJSON = JSON.stringify(updates.widgets);
      if (updates.settings) itemData.SettingsJSON = JSON.stringify(updates.settings);
      if (updates.globalFilters) itemData.GlobalFiltersJSON = JSON.stringify(updates.globalFilters);
      if (updates.isPublic !== undefined) itemData.IsPublic = updates.isPublic;
      if (updates.tags) itemData.Tags = updates.tags.join(';');

      await this.sp.web.lists.getByTitle(this.listName).items.getById(reportId).update(itemData);

      logger.info('ReportDefinitionService', `Updated report ID: ${reportId}`);
    } catch (error) {
      logger.error('ReportDefinitionService.updateReport', error);
      throw error;
    }
  }

  /**
   * Delete a report definition
   */
  public async deleteReport(reportId: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.listName).items.getById(reportId).delete();
      logger.info('ReportDefinitionService', `Deleted report ID: ${reportId}`);
    } catch (error) {
      logger.error('ReportDefinitionService.deleteReport', error);
      throw error;
    }
  }

  /**
   * Get reports created by current user
   */
  public async getMyReports(userId: number): Promise<IReportDefinition[]> {
    try {
      console.log(`[ReportDefinitionService] getMyReports - Loading reports for user ID ${userId} from ${this.listName}`);

      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select(
          'ID',
          'ReportName',
          'Description',
          'Category',
          'Status',
          'IsPublic',
          'IsTemplate',
          'Modified',
          'Author/Title'
        )
        .expand('Author')
        .filter(`AuthorId eq ${userId}`)
        .orderBy('Modified', false)();

      console.log(`[ReportDefinitionService] Retrieved ${items.length} reports for user ID ${userId}`);
      return items.map(item => this.mapItemToReportDefinition(item));
    } catch (error: any) {
      console.error(`[ReportDefinitionService] Error loading user reports:`, error?.message || error);
      logger.error('ReportDefinitionService.getMyReports', error);
      throw error;
    }
  }

  /**
   * Get public reports
   */
  public async getPublicReports(): Promise<IReportDefinition[]> {
    try {
      console.log(`[ReportDefinitionService] getPublicReports - Loading public reports from ${this.listName}`);

      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select(
          'ID',
          'ReportName',
          'Description',
          'Category',
          'Status',
          'IsPublic',
          'Modified',
          'Author/Title'
        )
        .expand('Author')
        .filter('IsPublic eq 1 and Status eq \'Active\'')
        .orderBy('Modified', false)();

      console.log(`[ReportDefinitionService] Retrieved ${items.length} public reports`);
      return items.map(item => this.mapItemToReportDefinition(item));
    } catch (error: any) {
      console.error(`[ReportDefinitionService] Error loading public reports:`, error?.message || error);
      logger.error('ReportDefinitionService.getPublicReports', error);
      throw error;
    }
  }

  /**
   * Get report templates
   */
  public async getTemplates(): Promise<IReportDefinition[]> {
    try {
      console.log(`[ReportDefinitionService] getTemplates - Loading templates from ${this.listName}`);

      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select(
          'ID',
          'ReportName',
          'Description',
          'Category',
          'WidgetsJSON',
          'LayoutConfig',
          'Modified'
        )
        .filter('IsTemplate eq 1 and Status eq \'Active\'')
        .orderBy('Category')();

      console.log(`[ReportDefinitionService] Retrieved ${items.length} templates`);
      return items.map(item => this.mapItemToReportDefinition(item));
    } catch (error: any) {
      console.error(`[ReportDefinitionService] Error loading templates:`, error?.message || error);
      logger.error('ReportDefinitionService.getTemplates', error);
      throw error;
    }
  }

  /**
   * Map SharePoint list item to IReportDefinition
   */
  private mapItemToReportDefinition(item: any): IReportDefinition {
    let layout: IReportLayout = {
      columns: 12,
      rows: 12,
      pageSize: 'A4',
      orientation: 'portrait',
      margins: {
        top: 20,
        right: 20,
        bottom: 20,
        left: 20
      }
    };

    if (item.LayoutConfig) {
      try {
        layout = JSON.parse(item.LayoutConfig);
      } catch (e) {
        console.warn('Failed to parse LayoutConfig', e);
      }
    }

    return {
      Id: item.ID,
      Title: item.ReportName || '',
      Description: item.Description || '',
      Category: item.Category as ReportCategory || ReportCategory.Custom,
      layout: layout,
      widgets: item.WidgetsJSON ? JSON.parse(item.WidgetsJSON) : [],
      settings: item.SettingsJSON ? JSON.parse(item.SettingsJSON) : {},
      globalFilters: item.GlobalFiltersJSON ? JSON.parse(item.GlobalFiltersJSON) : undefined,
      isPublic: item.IsPublic || false,
      tags: item.Tags ? item.Tags.split(';') : [],
      createdBy: item.Author?.ID,
      createdDate: new Date(item.Created),
      modifiedDate: new Date(item.Modified)
    };
  }
}
