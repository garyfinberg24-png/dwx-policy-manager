// @ts-nocheck
/**
 * ProcessAutomationService
 * Provides integration between SPFx and Power Automate flows for process orchestration.
 * This service can trigger Power Automate flows via HTTP and handle webhooks.
 */

import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { logger } from './LoggingService';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

/**
 * Power Automate Flow Configuration
 */
export interface IFlowConfig {
  id: string;
  name: string;
  triggerUrl: string;
  isEnabled: boolean;
  description?: string;
}

/**
 * Process Automation Trigger Payload
 */
export interface IProcessTriggerPayload {
  processId: number;
  processType: 'Joiner' | 'Mover' | 'Leaver';
  employeeName: string;
  employeeEmail: string;
  department: string;
  jobTitle: string;
  startDate: string;
  managerId?: number;
  managerEmail?: string;
  location?: string;
  totalTasks: number;
  triggeredBy: string;
  triggeredAt: string;
  additionalData?: Record<string, unknown>;
}

/**
 * Flow Trigger Response
 */
export interface IFlowTriggerResponse {
  success: boolean;
  flowRunId?: string;
  error?: string;
  timestamp: Date;
}

/**
 * Automation Log Entry
 */
export interface IAutomationLogEntry {
  Title: string;
  FlowName: string;
  ProcessId: string;
  EmployeeName?: string;
  ActionType: string;
  Status: 'Success' | 'Failed' | 'Pending';
  ErrorMessage?: string;
  ExecutionDate: string;
  FlowRunId?: string;
}

/**
 * ProcessAutomationService
 * Manages Power Automate integration for JML process orchestration
 */
export class ProcessAutomationService {
  private context: WebPartContext;
  private sp: SPFI;
  private flowConfigs: Map<string, IFlowConfig> = new Map();

  // Default flow configuration list
  private readonly FLOW_CONFIG_LIST = 'PM_PowerAutomateFlows';
  private readonly AUTOMATION_LOG_LIST = 'PM_AutomationLogs';

  constructor(context: WebPartContext, sp: SPFI) {
    this.context = context;
    this.sp = sp;
  }

  /**
   * Initialize the service and load flow configurations
   */
  public async initialize(): Promise<void> {
    try {
      await this.loadFlowConfigs();
      logger.info('ProcessAutomationService', 'Initialized with flow configurations');
    } catch (error) {
      logger.warn('ProcessAutomationService', 'Could not load flow configs - list may not exist', error);
      // Initialize with default/hardcoded configs if list doesn't exist
      this.initializeDefaultConfigs();
    }
  }

  /**
   * Load flow configurations from SharePoint list
   */
  private async loadFlowConfigs(): Promise<void> {
    const items = await this.sp.web.lists
      .getByTitle(this.FLOW_CONFIG_LIST)
      .items
      .filter('IsEnabled eq 1')
      .select('Id', 'Title', 'FlowId', 'TriggerUrl', 'IsEnabled', 'Description')();

    this.flowConfigs.clear();
    for (const item of items) {
      this.flowConfigs.set(item.FlowId, {
        id: item.FlowId,
        name: item.Title,
        triggerUrl: item.TriggerUrl,
        isEnabled: item.IsEnabled,
        description: item.Description
      });
    }

    logger.info('ProcessAutomationService', `Loaded ${this.flowConfigs.size} flow configurations`);
  }

  /**
   * Initialize default flow configurations (fallback)
   */
  private initializeDefaultConfigs(): void {
    // These URLs should be configured in environment or SharePoint list
    // Using placeholder URLs that would be replaced during deployment
    const defaultConfigs: IFlowConfig[] = [
      {
        id: 'PA-001',
        name: 'Joiner Process Orchestration',
        triggerUrl: '', // Configure during deployment
        isEnabled: false,
        description: 'Orchestrates welcome emails, calendar events, and notifications for new joiners'
      },
      {
        id: 'PA-002',
        name: 'Mover Process Orchestration',
        triggerUrl: '',
        isEnabled: false,
        description: 'Orchestrates notifications and IT requests for role/location changes'
      },
      {
        id: 'PA-003',
        name: 'Leaver Process Orchestration',
        triggerUrl: '',
        isEnabled: false,
        description: 'Orchestrates offboarding notifications, access revocation, and equipment collection'
      },
      {
        id: 'PA-004',
        name: 'Process Auto-Completion',
        triggerUrl: '',
        isEnabled: false,
        description: 'Automatically marks processes as completed when all tasks are done'
      }
    ];

    for (const config of defaultConfigs) {
      this.flowConfigs.set(config.id, config);
    }

    logger.info('ProcessAutomationService', 'Initialized with default configurations (flows not active)');
  }

  /**
   * Trigger the appropriate process orchestration flow based on process type
   */
  public async triggerProcessOrchestration(
    payload: IProcessTriggerPayload
  ): Promise<IFlowTriggerResponse> {
    // Determine which flow to trigger based on process type
    // Maps JML ProcessType enum values to Power Automate flow IDs
    let flowId: string;
    switch (payload.processType) {
      case 'Joiner':
        flowId = 'PA-001'; // Joiner Process Orchestration
        break;
      case 'Mover':
        flowId = 'PA-002'; // Mover Process Orchestration
        break;
      case 'Leaver':
        flowId = 'PA-003'; // Leaver Process Orchestration
        break;
      default:
        logger.warn('ProcessAutomationService', `Unknown process type: ${payload.processType}`);
        return {
          success: false,
          error: `Unknown process type: ${payload.processType}`,
          timestamp: new Date()
        };
    }

    return this.triggerFlow(flowId, payload);
  }

  /**
   * Trigger a specific Power Automate flow
   */
  public async triggerFlow(
    flowId: string,
    payload: IProcessTriggerPayload
  ): Promise<IFlowTriggerResponse> {
    const config = this.flowConfigs.get(flowId);

    if (!config) {
      logger.warn('ProcessAutomationService', `Flow configuration not found: ${flowId}`);
      return {
        success: false,
        error: `Flow configuration not found: ${flowId}`,
        timestamp: new Date()
      };
    }

    if (!config.isEnabled || !config.triggerUrl) {
      logger.info('ProcessAutomationService', `Flow ${flowId} is not enabled or configured`);
      // Log but don't fail - flows may not be deployed yet
      await this.logAutomation({
        Title: `Flow Trigger Skipped: ${config.name}`,
        FlowName: config.name,
        ProcessId: payload.processId.toString(),
        EmployeeName: payload.employeeName,
        ActionType: 'FlowTrigger',
        Status: 'Pending',
        ErrorMessage: 'Flow not enabled or trigger URL not configured',
        ExecutionDate: new Date().toISOString()
      });

      return {
        success: true, // Don't fail the process creation
        error: 'Flow not enabled - skipped',
        timestamp: new Date()
      };
    }

    try {
      logger.info('ProcessAutomationService', `Triggering flow: ${config.name}`, { processId: payload.processId });

      const httpClientOptions: IHttpClientOptions = {
        body: JSON.stringify({
          ...payload,
          flowId: flowId,
          source: 'JML-SPFx',
          version: '1.0.0'
        }),
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        }
      };

      const response: HttpClientResponse = await this.context.httpClient.post(
        config.triggerUrl,
        HttpClient.configurations.v1,
        httpClientOptions
      );

      if (response.ok) {
        const responseData = await response.json().catch(() => ({}));
        const flowRunId = responseData?.workflowRunId || responseData?.id;

        logger.info('ProcessAutomationService', `Flow triggered successfully: ${config.name}`, { flowRunId });

        await this.logAutomation({
          Title: `Flow Triggered: ${config.name}`,
          FlowName: config.name,
          ProcessId: payload.processId.toString(),
          EmployeeName: payload.employeeName,
          ActionType: 'FlowTrigger',
          Status: 'Success',
          FlowRunId: flowRunId,
          ExecutionDate: new Date().toISOString()
        });

        return {
          success: true,
          flowRunId: flowRunId,
          timestamp: new Date()
        };
      } else {
        const errorText = await response.text().catch(() => 'Unknown error');
        logger.error('ProcessAutomationService', `Flow trigger failed: ${response.status}`, { error: errorText });

        await this.logAutomation({
          Title: `Flow Trigger Failed: ${config.name}`,
          FlowName: config.name,
          ProcessId: payload.processId.toString(),
          EmployeeName: payload.employeeName,
          ActionType: 'FlowTrigger',
          Status: 'Failed',
          ErrorMessage: `HTTP ${response.status}: ${errorText.substring(0, 500)}`,
          ExecutionDate: new Date().toISOString()
        });

        return {
          success: false,
          error: `HTTP ${response.status}: ${errorText}`,
          timestamp: new Date()
        };
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ProcessAutomationService', 'Flow trigger exception', error);

      await this.logAutomation({
        Title: `Flow Trigger Error: ${config.name}`,
        FlowName: config.name,
        ProcessId: payload.processId.toString(),
        EmployeeName: payload.employeeName,
        ActionType: 'FlowTrigger',
        Status: 'Failed',
        ErrorMessage: errorMessage,
        ExecutionDate: new Date().toISOString()
      });

      return {
        success: false,
        error: errorMessage,
        timestamp: new Date()
      };
    }
  }

  /**
   * Log automation activity to SharePoint
   */
  private async logAutomation(entry: IAutomationLogEntry): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.AUTOMATION_LOG_LIST)
        .items.add(entry);
    } catch (error) {
      logger.warn('ProcessAutomationService', 'Failed to log automation activity', error);
      // Don't throw - logging failure shouldn't break the process
    }
  }

  /**
   * Get all configured flows
   */
  public getFlowConfigs(): IFlowConfig[] {
    return Array.from(this.flowConfigs.values());
  }

  /**
   * Get a specific flow configuration
   */
  public getFlowConfig(flowId: string): IFlowConfig | undefined {
    return this.flowConfigs.get(flowId);
  }

  /**
   * Update flow configuration (admin function)
   */
  public async updateFlowConfig(
    flowId: string,
    updates: Partial<IFlowConfig>
  ): Promise<void> {
    try {
      // Find the list item
      const items = await this.sp.web.lists
        .getByTitle(this.FLOW_CONFIG_LIST)
        .items
        .filter(`FlowId eq '${flowId}'`)
        .top(1)();

      if (items.length === 0) {
        throw new Error(`Flow configuration not found: ${flowId}`);
      }

      // Update the item
      await this.sp.web.lists
        .getByTitle(this.FLOW_CONFIG_LIST)
        .items.getById(items[0].Id)
        .update({
          Title: updates.name,
          TriggerUrl: updates.triggerUrl,
          IsEnabled: updates.isEnabled,
          Description: updates.description
        });

      // Reload configs
      await this.loadFlowConfigs();

      logger.info('ProcessAutomationService', `Flow config updated: ${flowId}`);
    } catch (error) {
      logger.error('ProcessAutomationService', 'Failed to update flow config', error);
      throw error;
    }
  }

  /**
   * Register a new flow configuration
   */
  public async registerFlow(config: IFlowConfig): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.FLOW_CONFIG_LIST)
        .items.add({
          Title: config.name,
          FlowId: config.id,
          TriggerUrl: config.triggerUrl,
          IsEnabled: config.isEnabled,
          Description: config.description
        });

      this.flowConfigs.set(config.id, config);
      logger.info('ProcessAutomationService', `Flow registered: ${config.id}`);
    } catch (error) {
      logger.error('ProcessAutomationService', 'Failed to register flow', error);
      throw error;
    }
  }

  /**
   * Test a flow trigger (dry run)
   */
  public async testFlowTrigger(flowId: string): Promise<IFlowTriggerResponse> {
    const config = this.flowConfigs.get(flowId);

    if (!config || !config.triggerUrl) {
      return {
        success: false,
        error: 'Flow not configured',
        timestamp: new Date()
      };
    }

    // Create test payload
    const testPayload: IProcessTriggerPayload = {
      processId: 0,
      processType: 'Joiner',
      employeeName: 'Test User',
      employeeEmail: 'test@example.com',
      department: 'Test Department',
      jobTitle: 'Test Position',
      startDate: new Date().toISOString(),
      totalTasks: 0,
      triggeredBy: 'System Test',
      triggeredAt: new Date().toISOString(),
      additionalData: {
        isTest: true
      }
    };

    return this.triggerFlow(flowId, testPayload);
  }
}

export default ProcessAutomationService;
