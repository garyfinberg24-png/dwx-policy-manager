// ============================================================================
// DWx Policy Manager - User Management Service
// CRUD for PM_Employees + SP Group management for role assignment
// ============================================================================

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/site-users/web';
import { IJMLEmployee } from '../models/IEntraUser';

// ============================================================================
// INTERFACES
// ============================================================================

export interface IEmployeeFilter {
  search?: string;
  role?: string;
  department?: string;
  status?: string;
}

export interface IEmployeePage {
  items: IJMLEmployee[];
  total: number;
}

export interface IRoleSummary {
  role: string;
  count: number;
  description: string;
}

export interface ISPGroupInfo {
  Id: number;
  Title: string;
  Description: string;
  OwnerTitle: string;
}

export interface ISPGroupMember {
  Id: number;
  Title: string;
  Email: string;
  LoginName: string;
  IsSiteAdmin: boolean;
}

// ============================================================================
// SERVICE
// ============================================================================

export class UserManagementService {
  private readonly sp: SPFI;
  private readonly EMPLOYEES_LIST = 'PM_Employees';

  private readonly EMPLOYEE_FIELDS = [
    'Id', 'Title', 'FirstName', 'LastName', 'Email', 'EmployeeNumber',
    'JobTitle', 'Department', 'Location', 'OfficePhone', 'MobilePhone',
    'ManagerEmail', 'Status', 'EmploymentType', 'CostCenter',
    'EntraObjectId', 'PMRole', 'LastSyncedAt', 'Notes', 'Created', 'Modified'
  ];

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==========================================================================
  // EMPLOYEE CRUD
  // ==========================================================================

  /**
   * Load employees with pagination and optional filters
   */
  public async getEmployees(page: number = 1, pageSize: number = 25, filters?: IEmployeeFilter): Promise<IEmployeePage> {
    try {
      const skip = (page - 1) * pageSize;
      let filterParts: string[] = [];

      if (filters?.role) {
        filterParts.push(`PMRole eq '${filters.role}'`);
      }
      if (filters?.department) {
        filterParts.push(`Department eq '${filters.department}'`);
      }
      if (filters?.status) {
        filterParts.push(`Status eq '${filters.status}'`);
      }
      if (filters?.search) {
        const term = filters.search.replace(/'/g, "''");
        filterParts.push(
          `(substringof('${term}',Title) or substringof('${term}',Email) or substringof('${term}',Department))`
        );
      }

      const filterStr = filterParts.length > 0 ? filterParts.join(' and ') : '';

      // Get total count
      let countQuery = this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('Id');
      if (filterStr) {
        countQuery = countQuery.filter(filterStr);
      }
      const allIds = await countQuery.top(5000)();
      const total = allIds.length;

      // Get page
      let itemsQuery = this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select(...this.EMPLOYEE_FIELDS)
        .orderBy('Title', true)
        .skip(skip)
        .top(pageSize);
      if (filterStr) {
        itemsQuery = itemsQuery.filter(filterStr);
      }
      const items: IJMLEmployee[] = await itemsQuery();

      return { items, total };
    } catch (err) {
      console.warn('UserManagementService.getEmployees failed:', err);
      return { items: [], total: 0 };
    }
  }

  /**
   * Search employees by name, email, or department
   */
  public async searchEmployees(query: string): Promise<IJMLEmployee[]> {
    try {
      const term = query.replace(/'/g, "''");
      const filter = `(substringof('${term}',Title) or substringof('${term}',Email) or substringof('${term}',Department))`;

      return await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select(...this.EMPLOYEE_FIELDS)
        .filter(filter)
        .orderBy('Title', true)
        .top(20)();
    } catch (err) {
      console.warn('UserManagementService.searchEmployees failed:', err);
      return [];
    }
  }

  /**
   * Get a single employee by ID
   */
  public async getEmployee(id: number): Promise<IJMLEmployee | null> {
    try {
      return await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .getById(id)
        .select(...this.EMPLOYEE_FIELDS)();
    } catch (err) {
      console.warn('UserManagementService.getEmployee failed:', err);
      return null;
    }
  }

  /**
   * Update PM role for a user
   */
  public async updateUserRole(employeeId: number, role: string): Promise<void> {
    await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
      .getById(employeeId)
      .update({ PMRole: role });
  }

  /**
   * Get role summary counts
   */
  public async getRoleSummary(): Promise<IRoleSummary[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('PMRole')
        .top(5000)();

      const counts: Record<string, number> = { Admin: 0, Manager: 0, Author: 0, User: 0 };
      for (const item of items) {
        const role = item.PMRole || 'User';
        counts[role] = (counts[role] || 0) + 1;
      }

      return [
        { role: 'Admin', count: counts.Admin, description: 'Full system access, all configuration' },
        { role: 'Manager', count: counts.Manager, description: 'Analytics, approvals, distribution, SLA' },
        { role: 'Author', count: counts.Author, description: 'Create policies, manage packs' },
        { role: 'User', count: counts.User, description: 'Browse, read, acknowledge policies' },
      ];
    } catch (err) {
      console.warn('UserManagementService.getRoleSummary failed:', err);
      return [
        { role: 'Admin', count: 0, description: 'Full system access, all configuration' },
        { role: 'Manager', count: 0, description: 'Analytics, approvals, distribution, SLA' },
        { role: 'Author', count: 0, description: 'Create policies, manage packs' },
        { role: 'User', count: 0, description: 'Browse, read, acknowledge policies' },
      ];
    }
  }

  /**
   * Get distinct department values from PM_Employees
   */
  public async getDepartments(): Promise<string[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('Department')
        .filter('Department ne null')
        .top(5000)();

      const unique = Array.from(new Set(items.map((i: any) => i.Department).filter(Boolean)));
      return unique.sort();
    } catch (err) {
      console.warn('UserManagementService.getDepartments failed:', err);
      return [];
    }
  }

  /**
   * Get distinct job titles from PM_Employees
   */
  public async getJobTitles(): Promise<string[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('JobTitle')
        .filter('JobTitle ne null')
        .top(5000)();

      const unique = Array.from(new Set(items.map((i: any) => i.JobTitle).filter(Boolean)));
      return unique.sort();
    } catch (err) {
      console.warn('UserManagementService.getJobTitles failed:', err);
      return [];
    }
  }

  /**
   * Get distinct locations from PM_Employees
   */
  public async getLocations(): Promise<string[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('Location')
        .filter('Location ne null')
        .top(5000)();

      const unique = Array.from(new Set(items.map((i: any) => i.Location).filter(Boolean)));
      return unique.sort();
    } catch (err) {
      console.warn('UserManagementService.getLocations failed:', err);
      return [];
    }
  }

  // ==========================================================================
  // SP GROUP MANAGEMENT
  // ==========================================================================

  /**
   * List SharePoint groups, optionally filtered by prefix
   */
  public async getSiteGroups(filterPrefix?: string): Promise<ISPGroupInfo[]> {
    try {
      const groups: any[] = await this.sp.web.siteGroups
        .select('Id', 'Title', 'Description', 'OwnerTitle')();

      let result = groups.map(g => ({
        Id: g.Id,
        Title: g.Title,
        Description: g.Description || '',
        OwnerTitle: g.OwnerTitle || '',
      }));

      if (filterPrefix) {
        result = result.filter(g => g.Title.startsWith(filterPrefix));
      }

      return result.sort((a, b) => a.Title.localeCompare(b.Title));
    } catch (err) {
      console.warn('UserManagementService.getSiteGroups failed:', err);
      return [];
    }
  }

  /**
   * Get members of a specific SP group
   */
  public async getGroupMembers(groupId: number): Promise<ISPGroupMember[]> {
    try {
      const users: any[] = await this.sp.web.siteGroups.getById(groupId).users
        .select('Id', 'Title', 'Email', 'LoginName', 'IsSiteAdmin')();

      return users.map(u => ({
        Id: u.Id,
        Title: u.Title,
        Email: u.Email || '',
        LoginName: u.LoginName,
        IsSiteAdmin: u.IsSiteAdmin || false,
      }));
    } catch (err) {
      console.warn('UserManagementService.getGroupMembers failed:', err);
      return [];
    }
  }

  /**
   * Add a user to a SP group by login name
   */
  public async addUserToGroup(groupId: number, loginName: string): Promise<void> {
    await this.sp.web.siteGroups.getById(groupId).users.add(loginName);
  }

  /**
   * Remove a user from a SP group
   */
  public async removeUserFromGroup(groupId: number, userId: number): Promise<void> {
    await this.sp.web.siteGroups.getById(groupId).users.removeById(userId);
  }

  /**
   * Create a new SP group
   */
  public async createGroup(name: string, description: string): Promise<ISPGroupInfo> {
    const result = await this.sp.web.siteGroups.add({
      Title: name,
      Description: description,
    });

    return {
      Id: result.data.Id,
      Title: result.data.Title,
      Description: result.data.Description || '',
      OwnerTitle: '',
    };
  }
}
