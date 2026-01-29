// @ts-nocheck
/**
 * EntraUserSyncService
 *
 * Comprehensive service for synchronizing users from Entra ID (Azure AD)
 * to the PM_Employees SharePoint list.
 *
 * Features:
 * - Full sync of all Entra users
 * - Incremental sync (delta queries)
 * - Single user sync
 * - Configurable field mapping
 * - Batch processing for large directories
 * - Sync logging and audit trail
 * - Error handling with retry logic
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import { graphfi, SPFx as GraphSPFx } from '@pnp/graph';
import { GraphFI } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import '@pnp/graph/users';
import '@pnp/graph/groups';
import '@pnp/graph/photos';

import {
  IEntraUser,
  IJMLEmployee,
  ISyncResult,
  ISyncSummary,
  ISyncConfig,
  SyncOperation,
  SyncStatus,
  DEFAULT_SYNC_CONFIG,
  DEFAULT_FIELD_MAPPINGS,
  IFieldMapping,
  EmployeeStatus
} from '../models/IEntraUser';

/**
 * Extended Entra user interface for delta queries
 * Includes @removed property indicating user was deleted
 */
interface IDeltaUser extends IEntraUser {
  '@removed'?: { reason: string };
}

/** List name for employee data */
const EMPLOYEES_LIST = 'PM_Employees';

/** List name for sync logs */
const SYNC_LOG_LIST = 'PM_Sync_Log';

/**
 * Service for synchronizing Entra ID users to JML Employees list
 */
export class EntraUserSyncService {
  private readonly sp: SPFI;
  private readonly graph: GraphFI;
  private readonly context: WebPartContext;
  private config: ISyncConfig;
  private fieldMappings: IFieldMapping[];

  /**
   * Creates an instance of EntraUserSyncService
   * @param context - SPFx webpart context
   * @param config - Optional sync configuration
   */
  constructor(context: WebPartContext, config?: Partial<ISyncConfig>) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
    this.graph = graphfi().using(GraphSPFx(context));
    this.config = { ...DEFAULT_SYNC_CONFIG, ...config };
    this.fieldMappings = DEFAULT_FIELD_MAPPINGS;
  }

  /**
   * Updates the sync configuration
   * @param config - Partial configuration to merge
   */
  public setConfig(config: Partial<ISyncConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * Sets custom field mappings
   * @param mappings - Array of field mappings
   */
  public setFieldMappings(mappings: IFieldMapping[]): void {
    this.fieldMappings = mappings;
  }

  /**
   * Performs a full sync of all Entra users to JML Employees
   * @returns Sync summary with results
   */
  public async syncAllUsers(): Promise<ISyncSummary> {
    const syncId = this.generateSyncId();
    const summary: ISyncSummary = {
      syncId,
      startedAt: new Date(),
      status: 'Running',
      totalProcessed: 0,
      added: 0,
      updated: 0,
      deactivated: 0,
      skipped: 0,
      errors: 0,
      results: [],
      errorDetails: []
    };

    try {
      // Log sync start
      await this.logSyncEvent(syncId, 'Started', 'Full sync initiated');

      // Fetch all Entra users
      const entraUsers = await this.fetchEntraUsers();
      summary.totalProcessed = entraUsers.length;

      // Fetch existing employees for matching
      const existingEmployees = await this.fetchExistingEmployees();

      // Process users in batches
      const batches = this.chunkArray(entraUsers, this.config.batchSize);

      for (const batch of batches) {
        const batchResults = await this.processBatch(batch, existingEmployees);
        summary.results.push(...batchResults);

        // Update counters
        for (const result of batchResults) {
          switch (result.operation) {
            case 'Added':
              summary.added++;
              break;
            case 'Updated':
              summary.updated++;
              break;
            case 'Deactivated':
              summary.deactivated++;
              break;
            case 'Skipped':
              summary.skipped++;
              break;
            case 'Error':
              summary.errors++;
              if (result.error) {
                summary.errorDetails?.push(`${result.userIdentifier}: ${result.error}`);
              }
              break;
          }
        }
      }

      // Handle deactivation of missing users
      if (this.config.deactivateMissing) {
        const deactivationResults = await this.deactivateMissingUsers(
          entraUsers,
          existingEmployees
        );
        summary.results.push(...deactivationResults);
        summary.deactivated += deactivationResults.filter(r => r.operation === 'Deactivated').length;
      }

      // Set final status
      summary.status = summary.errors > 0 ? 'CompletedWithErrors' : 'Completed';
      summary.completedAt = new Date();

      // Log sync completion
      await this.logSyncEvent(
        syncId,
        summary.status,
        `Completed: ${summary.added} added, ${summary.updated} updated, ${summary.errors} errors`
      );

      // Send notification if configured
      if (this.config.sendNotification) {
        await this.sendSyncNotification(summary);
      }

      return summary;
    } catch (error) {
      summary.status = 'Failed';
      summary.completedAt = new Date();
      summary.errorDetails?.push(`Fatal error: ${(error as Error).message}`);

      await this.logSyncEvent(syncId, 'Failed', (error as Error).message);

      throw error;
    }
  }

  /**
   * Syncs a single user from Entra ID
   * @param userIdentifier - Email or UPN of the user
   * @returns Sync result for the user
   */
  public async syncSingleUser(userIdentifier: string): Promise<ISyncResult> {
    try {
      // Fetch user from Entra
      const entraUser = await this.fetchEntraUser(userIdentifier);

      if (!entraUser) {
        return {
          userIdentifier,
          displayName: userIdentifier,
          operation: 'Error',
          success: false,
          error: 'User not found in Entra ID'
        };
      }

      // Check if user already exists in JML
      const existingEmployee = await this.findExistingEmployee(entraUser);

      // Prepare employee data
      const employeeData = this.mapEntraUserToEmployee(entraUser);

      if (existingEmployee) {
        if (this.config.updateExisting) {
          // Update existing employee
          await this.updateEmployee(existingEmployee.Id!, employeeData);
          return {
            userIdentifier: entraUser.mail || entraUser.userPrincipalName,
            displayName: entraUser.displayName,
            operation: 'Updated',
            success: true,
            itemId: existingEmployee.Id
          };
        } else {
          return {
            userIdentifier: entraUser.mail || entraUser.userPrincipalName,
            displayName: entraUser.displayName,
            operation: 'Skipped',
            success: true,
            itemId: existingEmployee.Id
          };
        }
      } else {
        // Add new employee
        const newItemId = await this.addEmployee(employeeData);
        return {
          userIdentifier: entraUser.mail || entraUser.userPrincipalName,
          displayName: entraUser.displayName,
          operation: 'Added',
          success: true,
          itemId: newItemId
        };
      }
    } catch (error) {
      return {
        userIdentifier,
        displayName: userIdentifier,
        operation: 'Error',
        success: false,
        error: (error as Error).message
      };
    }
  }

  /**
   * Syncs users from a specific Entra group
   * @param groupId - Entra group ID
   * @returns Sync summary
   */
  public async syncUsersFromGroup(groupId: string): Promise<ISyncSummary> {
    const syncId = this.generateSyncId();
    const summary: ISyncSummary = {
      syncId,
      startedAt: new Date(),
      status: 'Running',
      totalProcessed: 0,
      added: 0,
      updated: 0,
      deactivated: 0,
      skipped: 0,
      errors: 0,
      results: []
    };

    try {
      // Fetch group members using raw Graph call
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const accessToken = await tokenProvider.getToken('https://graph.microsoft.com');
      const response = await fetch(`https://graph.microsoft.com/v1.0/groups/${groupId}/members`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      const membersData = await response.json();
      const members = membersData.value || [];

      // Filter to users only
      const userMembers = members.filter(
        (m: { '@odata.type'?: string }) => m['@odata.type'] === '#microsoft.graph.user'
      );

      summary.totalProcessed = userMembers.length;

      // Fetch existing employees
      const existingEmployees = await this.fetchExistingEmployees();

      // Process each user
      for (const member of userMembers) {
        const entraUser = await this.fetchEntraUser(member.id);
        if (entraUser) {
          const result = await this.processUser(entraUser, existingEmployees);
          summary.results.push(result);

          if (result.operation === 'Added') summary.added++;
          else if (result.operation === 'Updated') summary.updated++;
          else if (result.operation === 'Skipped') summary.skipped++;
          else if (result.operation === 'Error') summary.errors++;
        }
      }

      summary.status = summary.errors > 0 ? 'CompletedWithErrors' : 'Completed';
      summary.completedAt = new Date();

      return summary;
    } catch (error) {
      summary.status = 'Failed';
      summary.completedAt = new Date();
      throw error;
    }
  }

  /**
   * Performs a delta sync - only syncs users changed since last sync
   * Uses Microsoft Graph delta queries for efficient incremental sync
   * @returns Sync summary with results
   */
  public async syncDelta(): Promise<ISyncSummary> {
    const syncId = this.generateSyncId();
    const summary: ISyncSummary = {
      syncId,
      startedAt: new Date(),
      status: 'Running',
      totalProcessed: 0,
      added: 0,
      updated: 0,
      deactivated: 0,
      skipped: 0,
      errors: 0,
      results: [],
      errorDetails: []
    };

    try {
      // Log sync start
      await this.logSyncEvent(syncId, 'Started', 'Delta sync initiated');

      // Get stored delta link from previous sync
      const deltaLink = await this.getDeltaLink();

      // Fetch changed users using delta query
      const { users: changedUsers, nextDeltaLink } = await this.fetchDeltaUsers(deltaLink);
      summary.totalProcessed = changedUsers.length;

      if (changedUsers.length === 0) {
        summary.status = 'Completed';
        summary.completedAt = new Date();
        await this.logSyncEvent(syncId, 'Completed', 'No changes detected since last sync');

        // Save the new delta link even if no changes
        if (nextDeltaLink) {
          await this.saveDeltaLink(nextDeltaLink);
        }

        return summary;
      }

      // Fetch existing employees for matching
      const existingEmployees = await this.fetchExistingEmployees();

      // Process changed users
      for (const user of changedUsers) {
        // Check if user was deleted (has @removed property)
        if ((user as IDeltaUser)['@removed']) {
          // Handle deleted user
          const result = await this.handleDeletedUser(user, existingEmployees);
          summary.results.push(result);
          if (result.operation === 'Deactivated') summary.deactivated++;
          else if (result.operation === 'Error') summary.errors++;
        } else {
          // Handle added/updated user
          const result = await this.processUser(user, existingEmployees);
          summary.results.push(result);

          switch (result.operation) {
            case 'Added':
              summary.added++;
              break;
            case 'Updated':
              summary.updated++;
              break;
            case 'Skipped':
              summary.skipped++;
              break;
            case 'Error':
              summary.errors++;
              if (result.error) {
                summary.errorDetails?.push(`${result.userIdentifier}: ${result.error}`);
              }
              break;
          }
        }
      }

      // Save the new delta link for next sync
      if (nextDeltaLink) {
        await this.saveDeltaLink(nextDeltaLink);
      }

      // Set final status
      summary.status = summary.errors > 0 ? 'CompletedWithErrors' : 'Completed';
      summary.completedAt = new Date();

      // Log sync completion
      await this.logSyncEvent(
        syncId,
        summary.status,
        `Delta sync completed: ${summary.added} added, ${summary.updated} updated, ${summary.deactivated} deactivated, ${summary.errors} errors`
      );

      // Send notification if configured
      if (this.config.sendNotification) {
        await this.sendSyncNotification(summary);
      }

      return summary;
    } catch (error) {
      summary.status = 'Failed';
      summary.completedAt = new Date();
      summary.errorDetails?.push(`Fatal error: ${(error as Error).message}`);

      await this.logSyncEvent(syncId, 'Failed', (error as Error).message);

      throw error;
    }
  }

  /**
   * Fetches users using Graph delta query
   * @param deltaLink - Previous delta link (null for initial sync)
   * @returns Changed users and next delta link
   */
  private async fetchDeltaUsers(deltaLink: string | null): Promise<{
    users: IEntraUser[];
    nextDeltaLink: string | null;
  }> {
    const selectFields = [
      'id',
      'userPrincipalName',
      'displayName',
      'givenName',
      'surname',
      'mail',
      'jobTitle',
      'department',
      'officeLocation',
      'businessPhones',
      'mobilePhone',
      'employeeId',
      'employeeType',
      'accountEnabled',
      'userType',
      'companyName'
    ];

    const users: IEntraUser[] = [];
    let nextDeltaLink: string | null = null;

    try {
      // Use the Graph client to make delta queries
      // Note: PnPjs doesn't have native delta support, so we use fetch
      const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
      let url: string;

      if (deltaLink) {
        // Use previous delta link
        url = deltaLink;
      } else {
        // Initial delta request
        url = `${graphBaseUrl}/users/delta?$select=${selectFields.join(',')}`;
      }

      // Get access token from context
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const accessToken = await tokenProvider.getToken('https://graph.microsoft.com');

      // Fetch all pages of delta results
      while (url) {
        const response = await fetch(url, {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        });

        if (!response.ok) {
          throw new Error(`Delta query failed: ${response.statusText}`);
        }

        const data = await response.json();

        // Add users from this page
        if (data.value) {
          for (const user of data.value) {
            // Apply filters
            if (this.shouldIncludeUser(user)) {
              users.push(user as IEntraUser);
            }
          }
        }

        // Check for next page or delta link
        if (data['@odata.nextLink']) {
          url = data['@odata.nextLink'];
        } else if (data['@odata.deltaLink']) {
          nextDeltaLink = data['@odata.deltaLink'];
          url = ''; // Exit loop
        } else {
          url = ''; // Exit loop
        }
      }

      return { users, nextDeltaLink };
    } catch (error) {
      console.error('Delta query error:', error);
      // Fall back to full sync if delta fails
      throw new Error(`Delta sync failed: ${(error as Error).message}. Consider running a full sync.`);
    }
  }

  /**
   * Checks if a user should be included based on config filters
   */
  private shouldIncludeUser(user: IEntraUser): boolean {
    // Filter by user type
    if (this.config.userTypeFilter && this.config.userTypeFilter.length > 0) {
      if (!this.config.userTypeFilter.includes(user.userType as 'Member' | 'Guest')) {
        return false;
      }
    }

    // Filter disabled users
    if (!this.config.includeDisabledUsers && user.accountEnabled === false) {
      return false;
    }

    // Filter by department
    if (this.config.departmentFilter && this.config.departmentFilter.length > 0) {
      if (!user.department || !this.config.departmentFilter.includes(user.department)) {
        return false;
      }
    }

    // Filter excluded users
    if (this.config.excludeUsers && this.config.excludeUsers.length > 0) {
      const excludeSet = new Set(this.config.excludeUsers.map(e => e.toLowerCase()));
      if (
        excludeSet.has(user.userPrincipalName.toLowerCase()) ||
        excludeSet.has((user.mail || '').toLowerCase())
      ) {
        return false;
      }
    }

    return true;
  }

  /**
   * Handles a deleted user from delta results
   */
  private async handleDeletedUser(
    user: IEntraUser,
    existingEmployees: Map<string, IJMLEmployee>
  ): Promise<ISyncResult> {
    try {
      // Find existing employee
      let existing = existingEmployees.get(user.id);
      if (!existing && user.mail) {
        existing = existingEmployees.get(user.mail.toLowerCase());
      }

      if (existing && existing.Status === 'Active') {
        // Deactivate the employee
        await this.updateEmployee(existing.Id!, { Status: 'Inactive' as EmployeeStatus });
        return {
          userIdentifier: user.mail || user.userPrincipalName || user.id,
          displayName: user.displayName || 'Unknown',
          operation: 'Deactivated',
          success: true,
          itemId: existing.Id
        };
      }

      return {
        userIdentifier: user.mail || user.userPrincipalName || user.id,
        displayName: user.displayName || 'Unknown',
        operation: 'Skipped',
        success: true
      };
    } catch (error) {
      return {
        userIdentifier: user.mail || user.userPrincipalName || user.id,
        displayName: user.displayName || 'Unknown',
        operation: 'Error',
        success: false,
        error: (error as Error).message
      };
    }
  }

  /**
   * Gets the stored delta link from previous sync
   */
  private async getDeltaLink(): Promise<string | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('PM_Sync_Config')
        .items
        .filter("ConfigType eq 'DeltaLink'")
        .top(1)();

      if (items.length > 0 && items[0].ConfigValue) {
        return items[0].ConfigValue;
      }
      return null;
    } catch {
      // Config list might not exist
      return null;
    }
  }

  /**
   * Saves the delta link for next sync
   */
  private async saveDeltaLink(deltaLink: string): Promise<void> {
    try {
      // Ensure config list exists
      await this.ensureSyncConfigList();

      // Check if delta link record exists
      const items = await this.sp.web.lists
        .getByTitle('PM_Sync_Config')
        .items
        .filter("ConfigType eq 'DeltaLink'")
        .top(1)();

      if (items.length > 0) {
        // Update existing
        await this.sp.web.lists
          .getByTitle('PM_Sync_Config')
          .items.getById(items[0].Id)
          .update({
            ConfigValue: deltaLink,
            Modified: new Date().toISOString()
          });
      } else {
        // Create new
        await this.sp.web.lists.getByTitle('PM_Sync_Config').items.add({
          Title: 'Delta Link',
          ConfigType: 'DeltaLink',
          ConfigValue: deltaLink
        });
      }
    } catch (error) {
      console.warn('Could not save delta link:', error);
    }
  }

  /**
   * Ensures the sync config list exists
   */
  private async ensureSyncConfigList(): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Sync_Config')();
    } catch {
      // Create the list if it doesn't exist
      await this.sp.web.lists.add('PM_Sync_Config', '', 100, false);
      const list = this.sp.web.lists.getByTitle('PM_Sync_Config');
      await list.fields.addText('ConfigType', { MaxLength: 50 });
      await list.fields.addMultilineText('ConfigValue', { NumberOfLines: 10 });
    }
  }

  /**
   * Resets the delta sync state (forces full sync on next delta call)
   */
  public async resetDeltaSync(): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('PM_Sync_Config')
        .items
        .filter("ConfigType eq 'DeltaLink'")
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists
          .getByTitle('PM_Sync_Config')
          .items.getById(items[0].Id)
          .delete();
      }
    } catch {
      // Ignore errors
    }
  }

  /**
   * Gets delta sync status
   */
  public async getDeltaSyncStatus(): Promise<{
    hasStoredDelta: boolean;
    lastDeltaSync: Date | null;
  }> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('PM_Sync_Config')
        .items
        .filter("ConfigType eq 'DeltaLink'")
        .select('Id', 'Modified')
        .top(1)();

      if (items.length > 0) {
        return {
          hasStoredDelta: true,
          lastDeltaSync: items[0].Modified ? new Date(items[0].Modified) : null
        };
      }
      return { hasStoredDelta: false, lastDeltaSync: null };
    } catch {
      return { hasStoredDelta: false, lastDeltaSync: null };
    }
  }

  /**
   * Gets the sync history/logs
   * @param count - Number of recent logs to retrieve
   * @returns Array of sync log entries
   */
  public async getSyncHistory(count: number = 10): Promise<ISyncLogEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(SYNC_LOG_LIST)
        .items
        .select('Id', 'Title', 'SyncId', 'Status', 'Message', 'Created')
        .orderBy('Created', false)
        .top(count)();

      return items.map((item: {
        Id: number;
        Title: string;
        SyncId: string;
        Status: string;
        Message: string;
        Created: string;
      }) => ({
        id: item.Id,
        syncId: item.SyncId,
        status: item.Status,
        message: item.Message,
        timestamp: new Date(item.Created)
      }));
    } catch {
      // Log list might not exist
      return [];
    }
  }

  // ============ Private Methods ============

  /**
   * Fetches all users from Entra ID with filtering
   */
  private async fetchEntraUsers(): Promise<IEntraUser[]> {
    const selectFields = [
      'id',
      'userPrincipalName',
      'displayName',
      'givenName',
      'surname',
      'mail',
      'jobTitle',
      'department',
      'officeLocation',
      'businessPhones',
      'mobilePhone',
      'employeeId',
      'employeeType',
      'accountEnabled',
      'userType',
      'companyName'
    ];

    let users = await this.graph.users
      .select(selectFields.join(','))
      .top(999)();

    // Apply filters
    if (this.config.userTypeFilter && this.config.userTypeFilter.length > 0) {
      users = users.filter((u: { userType?: string }) =>
        this.config.userTypeFilter!.includes(u.userType as 'Member' | 'Guest')
      );
    }

    if (!this.config.includeDisabledUsers) {
      users = users.filter((u: { accountEnabled?: boolean }) => u.accountEnabled === true);
    }

    if (this.config.departmentFilter && this.config.departmentFilter.length > 0) {
      users = users.filter((u: { department?: string }) =>
        u.department && this.config.departmentFilter!.includes(u.department)
      );
    }

    if (this.config.excludeUsers && this.config.excludeUsers.length > 0) {
      const excludeSet = new Set(this.config.excludeUsers.map(e => e.toLowerCase()));
      users = users.filter(
        (u: { userPrincipalName: string; mail?: string }) =>
          !excludeSet.has(u.userPrincipalName.toLowerCase()) &&
          !excludeSet.has((u.mail || '').toLowerCase())
      );
    }

    return users as IEntraUser[];
  }

  /**
   * Fetches a single user from Entra ID
   */
  private async fetchEntraUser(identifier: string): Promise<IEntraUser | null> {
    try {
      const user = await this.graph.users.getById(identifier)();
      return user as IEntraUser;
    } catch {
      return null;
    }
  }

  /**
   * Fetches existing employees from JML list
   */
  private async fetchExistingEmployees(): Promise<Map<string, IJMLEmployee>> {
    const employees = new Map<string, IJMLEmployee>();

    const items = await this.sp.web.lists
      .getByTitle(EMPLOYEES_LIST)
      .items
      .select(
        'Id',
        'Title',
        'Email',
        'EntraObjectId',
        'Status',
        'FirstName',
        'LastName',
        'Department',
        'JobTitle'
      )
      .top(5000)();

    for (const item of items) {
      const employee: IJMLEmployee = {
        Id: item.Id,
        Title: item.Title,
        Email: item.Email,
        EntraObjectId: item.EntraObjectId,
        Status: item.Status,
        FirstName: item.FirstName,
        LastName: item.LastName,
        Department: item.Department,
        JobTitle: item.JobTitle
      };

      // Index by EntraObjectId and Email for fast lookup
      if (item.EntraObjectId) {
        employees.set(item.EntraObjectId, employee);
      }
      if (item.Email) {
        employees.set(item.Email.toLowerCase(), employee);
      }
    }

    return employees;
  }

  /**
   * Finds existing employee matching an Entra user
   */
  private async findExistingEmployee(
    entraUser: IEntraUser
  ): Promise<IJMLEmployee | null> {
    const items = await this.sp.web.lists
      .getByTitle(EMPLOYEES_LIST)
      .items
      .filter(
        `EntraObjectId eq '${entraUser.id}' or Email eq '${entraUser.mail}'`
      )
      .top(1)();

    if (items.length > 0) {
      return items[0] as IJMLEmployee;
    }
    return null;
  }

  /**
   * Processes a batch of users
   */
  private async processBatch(
    users: IEntraUser[],
    existingEmployees: Map<string, IJMLEmployee>
  ): Promise<ISyncResult[]> {
    const results: ISyncResult[] = [];

    for (const user of users) {
      const result = await this.processUser(user, existingEmployees);
      results.push(result);
    }

    return results;
  }

  /**
   * Processes a single user sync
   */
  private async processUser(
    entraUser: IEntraUser,
    existingEmployees: Map<string, IJMLEmployee>
  ): Promise<ISyncResult> {
    try {
      const userIdentifier = entraUser.mail || entraUser.userPrincipalName;

      // Find existing employee
      let existing = existingEmployees.get(entraUser.id);
      if (!existing && entraUser.mail) {
        existing = existingEmployees.get(entraUser.mail.toLowerCase());
      }

      const employeeData = this.mapEntraUserToEmployee(entraUser);

      // Determine status based on Entra account status
      if (!entraUser.accountEnabled) {
        employeeData.Status = 'Inactive';
      }

      if (existing) {
        if (this.config.updateExisting) {
          await this.updateEmployee(existing.Id!, employeeData);
          return {
            userIdentifier,
            displayName: entraUser.displayName,
            operation: 'Updated',
            success: true,
            itemId: existing.Id
          };
        } else {
          return {
            userIdentifier,
            displayName: entraUser.displayName,
            operation: 'Skipped',
            success: true,
            itemId: existing.Id
          };
        }
      } else {
        const newId = await this.addEmployee(employeeData);
        return {
          userIdentifier,
          displayName: entraUser.displayName,
          operation: 'Added',
          success: true,
          itemId: newId
        };
      }
    } catch (error) {
      return {
        userIdentifier: entraUser.mail || entraUser.userPrincipalName,
        displayName: entraUser.displayName,
        operation: 'Error',
        success: false,
        error: (error as Error).message
      };
    }
  }

  /**
   * Maps Entra user to JML employee fields
   */
  private mapEntraUserToEmployee(entraUser: IEntraUser): Partial<IJMLEmployee> {
    const employee: Partial<IJMLEmployee> = {
      Status: 'Active',
      LastSyncedAt: new Date()
    };

    for (const mapping of this.fieldMappings) {
      if (mapping.enabled) {
        const value = entraUser[mapping.entraField];
        if (value !== undefined && value !== null) {
          // Handle phone array
          if (mapping.entraField === 'businessPhones' && Array.isArray(value)) {
            (employee as Record<string, unknown>)[mapping.jmlField] = value[0] || '';
          } else {
            (employee as Record<string, unknown>)[mapping.jmlField] = value;
          }
        }
      }
    }

    return employee;
  }

  /**
   * Adds a new employee to the list
   */
  private async addEmployee(data: Partial<IJMLEmployee>): Promise<number> {
    const result = await this.sp.web.lists
      .getByTitle(EMPLOYEES_LIST)
      .items.add(this.prepareListItemData(data));

    return (result.data as { Id: number }).Id;
  }

  /**
   * Updates an existing employee
   */
  private async updateEmployee(
    id: number,
    data: Partial<IJMLEmployee>
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(EMPLOYEES_LIST)
      .items.getById(id)
      .update(this.prepareListItemData(data));
  }

  /**
   * Prepares data for SharePoint list item
   */
  private prepareListItemData(
    data: Partial<IJMLEmployee>
  ): Record<string, unknown> {
    const listData: Record<string, unknown> = {};

    // Map fields, handling dates specially
    for (const [key, value] of Object.entries(data)) {
      if (value instanceof Date) {
        listData[key] = value.toISOString();
      } else if (value !== undefined && value !== null) {
        listData[key] = value;
      }
    }

    return listData;
  }

  /**
   * Deactivates employees not found in Entra
   */
  private async deactivateMissingUsers(
    entraUsers: IEntraUser[],
    existingEmployees: Map<string, IJMLEmployee>
  ): Promise<ISyncResult[]> {
    const results: ISyncResult[] = [];
    const entraIds = new Set(entraUsers.map(u => u.id));
    const entraEmails = new Set(
      entraUsers.map(u => (u.mail || '').toLowerCase()).filter(e => e)
    );

    const processedIds = new Set<number>();

    for (const [, employee] of Array.from(existingEmployees.entries())) {
      if (!employee.Id || processedIds.has(employee.Id)) continue;
      processedIds.add(employee.Id);

      const inEntra =
        (employee.EntraObjectId && entraIds.has(employee.EntraObjectId)) ||
        (employee.Email && entraEmails.has(employee.Email.toLowerCase()));

      if (!inEntra && employee.Status === 'Active') {
        try {
          await this.updateEmployee(employee.Id, { Status: 'Inactive' as EmployeeStatus });
          results.push({
            userIdentifier: employee.Email || employee.Title,
            displayName: employee.Title,
            operation: 'Deactivated',
            success: true,
            itemId: employee.Id
          });
        } catch (error) {
          results.push({
            userIdentifier: employee.Email || employee.Title,
            displayName: employee.Title,
            operation: 'Error',
            success: false,
            error: (error as Error).message
          });
        }
      }
    }

    return results;
  }

  /**
   * Logs a sync event to the sync log list
   */
  private async logSyncEvent(
    syncId: string,
    status: string,
    message: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(SYNC_LOG_LIST).items.add({
        Title: `Sync ${syncId}`,
        SyncId: syncId,
        Status: status,
        Message: message
      });
    } catch {
      // Sync log list might not exist - fail silently
      console.warn('Could not log sync event - PM_Sync_Log list may not exist');
    }
  }

  /**
   * Sends notification email about sync completion
   */
  private async sendSyncNotification(summary: ISyncSummary): Promise<void> {
    if (!this.config.notificationRecipients?.length) return;

    // Use Graph to send email via fetch
    const mailBody = {
      message: {
        subject: `JML User Sync ${summary.status} - ${summary.syncId}`,
        body: {
          contentType: 'HTML',
          content: this.formatSyncSummaryHtml(summary)
        },
        toRecipients: this.config.notificationRecipients.map(email => ({
          emailAddress: { address: email }
        }))
      }
    };

    try {
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const accessToken = await tokenProvider.getToken('https://graph.microsoft.com');

      await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(mailBody)
      });
    } catch (error) {
      console.error('Failed to send sync notification:', error);
    }
  }

  /**
   * Formats sync summary as HTML for email
   */
  private formatSyncSummaryHtml(summary: ISyncSummary): string {
    return `
      <h2>JML User Sync Summary</h2>
      <p><strong>Sync ID:</strong> ${summary.syncId}</p>
      <p><strong>Status:</strong> ${summary.status}</p>
      <p><strong>Started:</strong> ${summary.startedAt.toISOString()}</p>
      <p><strong>Completed:</strong> ${summary.completedAt?.toISOString() || 'N/A'}</p>
      <hr/>
      <h3>Results</h3>
      <ul>
        <li>Total Processed: ${summary.totalProcessed}</li>
        <li>Added: ${summary.added}</li>
        <li>Updated: ${summary.updated}</li>
        <li>Deactivated: ${summary.deactivated}</li>
        <li>Skipped: ${summary.skipped}</li>
        <li>Errors: ${summary.errors}</li>
      </ul>
      ${
        summary.errorDetails?.length
          ? `<h3>Errors</h3><ul>${summary.errorDetails
              .map(e => `<li>${e}</li>`)
              .join('')}</ul>`
          : ''
      }
    `;
  }

  /**
   * Generates a unique sync ID
   */
  private generateSyncId(): string {
    const timestamp = new Date().toISOString().replace(/[^0-9]/g, '').slice(0, 14);
    const random = Math.random().toString(36).substring(2, 8);
    return `SYNC-${timestamp}-${random}`;
  }

  /**
   * Chunks an array into batches
   */
  private chunkArray<T>(array: T[], size: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  }
}

/**
 * Sync log entry interface
 */
export interface ISyncLogEntry {
  id: number;
  syncId: string;
  status: string;
  message: string;
  timestamp: Date;
}

export default EntraUserSyncService;
