// @ts-nocheck
// GraphService - Microsoft Graph API access
// For user profiles, Teams, Planner, Exchange integrations

import { GraphFI } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/groups';
import '@pnp/graph/teams';
import '@pnp/graph/planner';
import {
  IEntraIDEmployee,
  ITeamsChannel,
  ITeamsChannelRequest,
  IPlannerPlan,
  IPlannerTask,
  IIntegrationResponse
} from '../models/IIntegration';
import { logger } from './LoggingService';
import { ValidationUtils } from '../utils/ValidationUtils';

export class GraphService {
  private graph: GraphFI;

  constructor(graph: GraphFI) {
    this.graph = graph;
  }

  /**
   * Get user by email (Enhanced for Entra ID)
   */
  public async getUserByEmail(email: string): Promise<IIntegrationResponse<IEntraIDEmployee>> {
    try {
      const user = await this.graph.users.getById(email).select(
        'id',
        'userPrincipalName',
        'displayName',
        'givenName',
        'surname',
        'mail',
        'mobilePhone',
        'jobTitle',
        'department',
        'officeLocation',
        'employeeId',
        'companyName'
      )();

      // Get manager
      let manager;
      try {
        manager = await this.graph.users.getById(email).manager.select('id', 'displayName', 'mail')();
      } catch {
        // Manager not set
      }

      const employee: IEntraIDEmployee = {
        id: user.id,
        userPrincipalName: user.userPrincipalName,
        displayName: user.displayName,
        givenName: user.givenName,
        surname: user.surname,
        mail: user.mail,
        mobilePhone: user.mobilePhone,
        jobTitle: user.jobTitle,
        department: user.department,
        officeLocation: user.officeLocation,
        employeeId: user.employeeId,
        companyName: user.companyName,
        manager: manager ? {
          id: manager.id,
          displayName: manager.displayName,
          mail: manager.mail
        } : undefined
      };

      return {
        success: true,
        data: employee,
        timestamp: new Date()
      };
    } catch (error) {
      logger.error('GraphService', `Error fetching user ${email}:`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to get user',
        timestamp: new Date()
      };
    }
  }

  /**
   * Search users
   */
  public async searchUsers(searchTerm: string): Promise<any[]> {
    try {
      const users = await this.graph.users
        .filter(`startswith(displayName,'${ValidationUtils.sanitizeForOData(searchTerm)}') or startswith(mail,'${ValidationUtils.sanitizeForOData(searchTerm)}')`)
        .top(10)();
      return users;
    } catch (error) {
      logger.error('GraphService', 'Error searching users:', error);
      throw error;
    }
  }

  /**
   * Get current user profile
   */
  public async getCurrentUserProfile(): Promise<any> {
    try {
      const user = await this.graph.me();
      return user;
    } catch (error) {
      logger.error('GraphService', 'Error fetching current user profile:', error);
      throw error;
    }
  }

  /**
   * Get user's manager
   */
  public async getUserManager(userId: string): Promise<any> {
    try {
      const manager = await this.graph.users.getById(userId).manager();
      return manager;
    } catch (error) {
      logger.error('GraphService', `Error fetching manager for ${userId}:`, error);
      return null;
    }
  }

  /**
   * Get user's direct reports
   */
  public async getUserDirectReports(userId: string): Promise<any[]> {
    try {
      const reports = await this.graph.users.getById(userId).directReports();
      return reports;
    } catch (error) {
      logger.error('GraphService', `Error fetching direct reports for ${userId}:`, error);
      return [];
    }
  }

  /**
   * Send email
   */
  public async sendEmail(
    to: string,
    subject: string,
    body: string,
    attachments?: Array<{ name: string; contentBytes: string }>
  ): Promise<void> {
    try {
      const message: any = {
        message: {
          subject,
          body: {
            contentType: 'HTML',
            content: body
          },
          toRecipients: [{
            emailAddress: {
              address: to
            }
          }]
        }
      };

      if (attachments && attachments.length > 0) {
        message.message.attachments = attachments.map(att => ({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: att.name,
          contentBytes: att.contentBytes
        }));
      }

      await (this.graph.me as any).sendMail(message);
    } catch (error) {
      logger.error('GraphService', 'Error sending email:', error);
      throw error;
    }
  }

  /**
   * Create Teams channel
   */
  public async createTeamsChannel(request: ITeamsChannelRequest): Promise<IIntegrationResponse<ITeamsChannel>> {
    try {
      const channel = await (this.graph.teams.getById(request.teamId) as any).channels.add({
        displayName: request.displayName,
        description: request.description || '',
        membershipType: request.membershipType || 'standard'
      });

      return {
        success: true,
        data: channel as ITeamsChannel,
        timestamp: new Date()
      };
    } catch (error) {
      logger.error('GraphService', 'Failed to create Teams channel:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create channel',
        timestamp: new Date()
      };
    }
  }

  /**
   * Create Planner plan
   */
  public async createPlannerPlan(groupId: string, title: string): Promise<IIntegrationResponse<IPlannerPlan>> {
    try {
      const plan = await (this.graph.planner as any).plans.add({
        owner: groupId,
        title: title
      });

      return {
        success: true,
        data: plan as IPlannerPlan,
        timestamp: new Date()
      };
    } catch (error) {
      logger.error('GraphService', 'Failed to create Planner plan:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create plan',
        timestamp: new Date()
      };
    }
  }

  /**
   * Create Planner task
   */
  public async createPlannerTask(task: IPlannerTask): Promise<IIntegrationResponse<IPlannerTask>> {
    try {
      const taskData: any = {
        planId: task.planId,
        title: task.title,
        percentComplete: task.percentComplete
      };

      if (task.bucketId) {
        taskData.bucketId = task.bucketId;
      }
      if (task.startDateTime) {
        taskData.startDateTime = task.startDateTime.toISOString();
      }
      if (task.dueDateTime) {
        taskData.dueDateTime = task.dueDateTime.toISOString();
      }
      if (task.assignments) {
        taskData.assignments = task.assignments;
      }
      if (task.priority !== undefined) {
        taskData.priority = task.priority;
      }

      const createdTask = await (this.graph.planner as any).tasks.add(taskData);

      return {
        success: true,
        data: createdTask as IPlannerTask,
        timestamp: new Date()
      };
    } catch (error) {
      logger.error('GraphService', 'Failed to create Planner task:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create task',
        timestamp: new Date()
      };
    }
  }

  /**
   * Update Planner task
   */
  public async updatePlannerTask(taskId: string, updates: Partial<IPlannerTask>): Promise<IIntegrationResponse<IPlannerTask>> {
    try {
      const updateData: any = {};

      if (updates.title) {
        updateData.title = updates.title;
      }
      if (updates.percentComplete !== undefined) {
        updateData.percentComplete = updates.percentComplete;
      }
      if (updates.startDateTime) {
        updateData.startDateTime = updates.startDateTime.toISOString();
      }
      if (updates.dueDateTime) {
        updateData.dueDateTime = updates.dueDateTime.toISOString();
      }
      if (updates.assignments) {
        updateData.assignments = updates.assignments;
      }
      if (updates.priority !== undefined) {
        updateData.priority = updates.priority;
      }

      const updatedTask = await (this.graph.planner as any).tasks.getById(taskId).update(updateData);

      return {
        success: true,
        data: updatedTask as IPlannerTask,
        timestamp: new Date()
      };
    } catch (error) {
      logger.error('GraphService', 'Failed to update Planner task:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update task',
        timestamp: new Date()
      };
    }
  }
}

