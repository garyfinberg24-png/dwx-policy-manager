// @ts-nocheck
// IntegrationService - Integration Hub for M365 and external systems
// Orchestrates integrations with Entra ID, Teams, Planner, Exchange, Power Automate, and HR systems

import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { GraphService } from './GraphService';
import {
  IEntraIDEmployee,
  ITeamsChannel,
  ITeamsChannelRequest,
  IPlannerTask,
  IPlannerPlan,
  IPowerAutomateTrigger,
  IIntegrationResponse,
  IIntegrationConfig,
  IIntegrationLog,
  IntegrationType,
  IntegrationStatus
} from '../models/IIntegration';
import { IJmlProcess } from '../models/IJmlProcess';
import { IJmlTask } from '../models/IJmlTask';
import { logger } from './LoggingService';

export interface IEmployee {
  id: string;
  employeeId?: string;
  displayName: string;
  email: string;
  jobTitle?: string;
  department?: string;
  manager?: {
    id: string;
    displayName: string;
    email: string;
  };
  officeLocation?: string;
  mobilePhone?: string;
}

export class IntegrationService {
  private sp: SPFI;
  private graph: GraphFI;
  private graphService: GraphService;
  private integrationConfigs: Map<IntegrationType, IIntegrationConfig> = new Map();
  private initialized: boolean = false;

  constructor(sp: SPFI, graph: GraphFI) {
    this.sp = sp;
    this.graph = graph;
    this.graphService = new GraphService(graph);
  }

  /**
   * Initialize service and load integration configurations
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      // Load integration configurations from SharePoint list
      const configs = await this.sp.web.lists
        .getByTitle('PM_IntegrationConfigs')
        .items
        .select('Id', 'Title', 'IntegrationType', 'Status', 'IsEnabled', 'Configuration')
        .filter('IsEnabled eq true')();

      configs.forEach(config => {
        this.integrationConfigs.set(config.IntegrationType, config as IIntegrationConfig);
      });

      this.initialized = true;
    } catch (error) {
      logger.error('IntegrationService', 'Failed to initialize IntegrationService:', error);
      // Continue without configs - individual methods will handle missing configs
      this.initialized = true;
    }
  }

  /**
   * Sync with Active Directory/Entra ID to auto-populate employee info
   * @param employeeId - Employee ID or email
   * @returns Employee information from Entra ID
   */
  public async syncWithAD(employeeId: string): Promise<IEmployee> {
    await this.ensureInitialized();

    const logEntry = {
      integrationType: IntegrationType.EntraID,
      action: 'syncWithAD',
      entityId: employeeId,
      startTime: new Date()
    };

    try {
      // Fetch user from Entra ID via Graph API
      const response = await this.graphService.getUserByEmail(employeeId);

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to fetch user from Entra ID');
      }

      const entraUser = response.data;

      // Map to IEmployee interface
      const employee: IEmployee = {
        id: entraUser.id,
        employeeId: entraUser.employeeId,
        displayName: entraUser.displayName,
        email: entraUser.mail,
        jobTitle: entraUser.jobTitle,
        department: entraUser.department,
        officeLocation: entraUser.officeLocation,
        mobilePhone: entraUser.mobilePhone,
        manager: entraUser.manager ? {
          id: entraUser.manager.id,
          displayName: entraUser.manager.displayName,
          email: entraUser.manager.mail
        } : undefined
      };

      // Log successful sync
      await this.logIntegration({
        ...logEntry,
        status: 'Success',
        responseData: JSON.stringify(employee),
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      return employee;
    } catch (error) {
      // Log failed sync
      await this.logIntegration({
        ...logEntry,
        status: 'Failed',
        errorMessage: error instanceof Error ? error.message : 'Unknown error',
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      throw error;
    }
  }

  /**
   * Create Microsoft Teams channel for new joiners
   * @param processId - JML Process ID
   * @returns Channel ID
   */
  public async createTeamsChannel(processId: number): Promise<string> {
    await this.ensureInitialized();

    const logEntry = {
      integrationType: IntegrationType.MicrosoftTeams,
      action: 'createTeamsChannel',
      processId,
      startTime: new Date()
    };

    try {
      // Get process details
      const process = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(processId)
        .select('Id', 'Title', 'EmployeeName', 'ProcessType', 'Department')();

      // Get Teams configuration
      const teamsConfig = this.integrationConfigs.get(IntegrationType.MicrosoftTeams);
      if (!teamsConfig || !teamsConfig.Configuration) {
        throw new Error('Teams integration not configured');
      }

      const config = JSON.parse(teamsConfig.Configuration);
      const teamId = config.defaultTeamId || config.teamId;

      if (!teamId) {
        throw new Error('Default Team ID not configured');
      }

      // Create channel name based on process
      const channelName = `${process.ProcessType} - ${process.EmployeeName}`;
      const description = `JML process channel for ${process.EmployeeName} (${process.Department})`;

      // Create Teams channel
      const channelRequest: ITeamsChannelRequest = {
        teamId,
        displayName: channelName,
        description,
        membershipType: 'standard'
      };

      const response = await this.graphService.createTeamsChannel(channelRequest);

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to create Teams channel');
      }

      const channelId = response.data.id;

      // Update process with channel link
      await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(processId)
        .update({
          CustomFields: JSON.stringify({
            teamsChannelId: channelId,
            teamsChannelUrl: response.data.webUrl
          })
        });

      // Log successful creation
      await this.logIntegration({
        ...logEntry,
        status: 'Success',
        responseData: JSON.stringify({ channelId, channelName }),
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      return channelId;
    } catch (error) {
      // Log failed creation
      await this.logIntegration({
        ...logEntry,
        status: 'Failed',
        errorMessage: error instanceof Error ? error.message : 'Unknown error',
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      throw error;
    }
  }

  /**
   * Sync JML tasks to Microsoft Planner
   * @param tasks - Array of JML tasks to assign
   */
  public async assignPlannerTasks(tasks: IJmlTask[]): Promise<void> {
    await this.ensureInitialized();

    const logEntry = {
      integrationType: IntegrationType.Planner,
      action: 'assignPlannerTasks',
      startTime: new Date()
    };

    try {
      // Get Planner configuration
      const plannerConfig = this.integrationConfigs.get(IntegrationType.Planner);
      if (!plannerConfig || !plannerConfig.Configuration) {
        throw new Error('Planner integration not configured');
      }

      const config = JSON.parse(plannerConfig.Configuration);
      const planId = config.defaultPlanId || config.planId;

      if (!planId) {
        throw new Error('Default Plan ID not configured');
      }

      const results: Array<{ taskId: number; plannerId?: string; error?: string }> = [];

      // Create Planner task for each JML task
      for (const task of tasks) {
        try {
          const plannerTask: IPlannerTask = {
            planId,
            title: task.Title,
            percentComplete: 0,
            priority: this.mapPriorityToPlanner(task.Priority),
            dueDateTime: task.SLAHours ? this.calculateDueDate(task.SLAHours) : undefined
          };

          const response = await this.graphService.createPlannerTask(plannerTask);

          if (response.success && response.data) {
            results.push({
              taskId: task.Id,
              plannerId: response.data.id
            });

            // Update JML task with Planner task ID
            await this.sp.web.lists
              .getByTitle('PM_Tasks')
              .items
              .getById(task.Id)
              .update({
                SystemUrl: `https://tasks.office.com/task/${response.data.id}`
              });
          } else {
            results.push({
              taskId: task.Id,
              error: response.error || 'Unknown error'
            });
          }
        } catch (error) {
          results.push({
            taskId: task.Id,
            error: error instanceof Error ? error.message : 'Unknown error'
          });
        }
      }

      const successCount = results.filter(r => !r.error).length;
      const failedCount = results.filter(r => r.error).length;

      // Log sync results
      await this.logIntegration({
        ...logEntry,
        status: failedCount === 0 ? 'Success' : (successCount > 0 ? 'Warning' : 'Failed'),
        responseData: JSON.stringify({
          total: tasks.length,
          success: successCount,
          failed: failedCount,
          results
        }),
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      if (failedCount > 0 && successCount === 0) {
        throw new Error(`Failed to sync all tasks to Planner. ${failedCount} failed.`);
      }
    } catch (error) {
      // Log failed sync
      await this.logIntegration({
        ...logEntry,
        status: 'Failed',
        errorMessage: error instanceof Error ? error.message : 'Unknown error',
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      throw error;
    }
  }

  /**
   * Trigger Power Automate flow with custom data
   * @param flowId - Power Automate Flow ID
   * @param data - Data to pass to the flow
   */
  public async triggerPowerAutomate(flowId: string, data: any): Promise<void> {
    await this.ensureInitialized();

    const logEntry = {
      integrationType: IntegrationType.PowerAutomate,
      action: 'triggerPowerAutomate',
      entityId: flowId,
      startTime: new Date()
    };

    try {
      // Get Power Automate configuration
      const paConfig = this.integrationConfigs.get(IntegrationType.PowerAutomate);
      if (!paConfig || !paConfig.EndpointUrl) {
        throw new Error('Power Automate integration not configured');
      }

      // Build webhook URL (HTTP trigger endpoint)
      const webhookUrl = paConfig.EndpointUrl.replace('{flowId}', flowId);

      // Prepare request payload
      const payload = {
        ...data,
        timestamp: new Date().toISOString(),
        source: 'PM_SPFx'
      };

      // Trigger the flow via HTTP POST
      const response = await fetch(webhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          ...(paConfig.ApiKey && { 'Authorization': `Bearer ${paConfig.ApiKey}` })
        },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Power Automate flow trigger failed: ${response.status} - ${errorText}`);
      }

      const responseData = await response.json();

      // Log successful trigger
      await this.logIntegration({
        ...logEntry,
        status: 'Success',
        requestData: JSON.stringify(payload),
        responseData: JSON.stringify(responseData),
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });
    } catch (error) {
      // Log failed trigger
      await this.logIntegration({
        ...logEntry,
        status: 'Failed',
        requestData: JSON.stringify(data),
        errorMessage: error instanceof Error ? error.message : 'Unknown error',
        executionTime: new Date().getTime() - logEntry.startTime.getTime()
      });

      throw error;
    }
  }

  /**
   * Get integration configuration
   */
  public getIntegrationConfig(type: IntegrationType): IIntegrationConfig | undefined {
    return this.integrationConfigs.get(type);
  }

  /**
   * Check integration health status
   */
  public async checkIntegrationHealth(type: IntegrationType): Promise<boolean> {
    const config = this.integrationConfigs.get(type);
    if (!config || !config.IsEnabled) {
      return false;
    }

    return config.Status === IntegrationStatus.Active;
  }

  /**
   * Log integration event to SharePoint
   */
  private async logIntegration(log: {
    integrationType: IntegrationType;
    action: string;
    status: 'Success' | 'Failed' | 'Warning';
    processId?: number;
    entityId?: string;
    requestData?: string;
    responseData?: string;
    errorMessage?: string;
    executionTime?: number;
  }): Promise<void> {
    try {
      const config = this.integrationConfigs.get(log.integrationType);

      await this.sp.web.lists
        .getByTitle('PM_IntegrationLogs')
        .items
        .add({
          Title: `${log.integrationType} - ${log.action}`,
          IntegrationConfigId: config?.Id,
          IntegrationType: log.integrationType,
          ProcessID: log.processId,
          Action: log.action,
          Status: log.status,
          RequestData: log.requestData,
          ResponseData: log.responseData,
          ErrorMessage: log.errorMessage,
          ExecutionTime: log.executionTime
        });
    } catch (error) {
      // Don't fail the operation if logging fails
      logger.error('IntegrationService', 'Failed to log integration event:', error);
    }
  }

  /**
   * Ensure service is initialized
   */
  private async ensureInitialized(): Promise<void> {
    if (!this.initialized) {
      await this.initialize();
    }
  }

  /**
   * Map JML priority to Planner priority (0-10)
   */
  private mapPriorityToPlanner(priority: string): number {
    switch (priority) {
      case 'High':
        return 1;
      case 'Medium':
        return 5;
      case 'Low':
        return 9;
      default:
        return 5;
    }
  }

  /**
   * Calculate due date based on SLA hours
   */
  private calculateDueDate(slaHours: number): Date {
    const dueDate = new Date();
    dueDate.setHours(dueDate.getHours() + slaHours);
    return dueDate;
  }
}
