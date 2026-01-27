// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/site-users/web';

/**
 * User roles for JML solution role-based access
 */
export enum UserRole {
  Employee = 'Employee',
  Manager = 'Manager',
  HRAdmin = 'HRAdmin',
  Recruiter = 'Recruiter',
  ITAdmin = 'ITAdmin',
  Executive = 'Executive',
  SiteAdmin = 'SiteAdmin',
  ProcurementManager = 'ProcurementManager',
  SkillsManager = 'SkillsManager',
  ContractManager = 'ContractManager',
  FinanceAdmin = 'FinanceAdmin'
}

/**
 * SharePoint group to role mapping configuration
 */
export interface IRoleMapping {
  groupName: string;
  role: UserRole;
}

/**
 * Default SharePoint group to role mappings
 * These can be customized in the Launchpad property pane
 */
export const DEFAULT_ROLE_MAPPINGS: IRoleMapping[] = [
  // Employee role - default for all users
  { groupName: 'All Users', role: UserRole.Employee },
  { groupName: 'Company Members', role: UserRole.Employee },

  // Manager role
  { groupName: 'Managers', role: UserRole.Manager },
  { groupName: 'JML Managers', role: UserRole.Manager },
  { groupName: 'People Managers', role: UserRole.Manager },

  // HR Admin role
  { groupName: 'HR Administrators', role: UserRole.HRAdmin },
  { groupName: 'HR Team', role: UserRole.HRAdmin },
  { groupName: 'JML HR Admins', role: UserRole.HRAdmin },

  // Recruiter role
  { groupName: 'Recruiters', role: UserRole.Recruiter },
  { groupName: 'Talent Acquisition', role: UserRole.Recruiter },
  { groupName: 'JML Recruiters', role: UserRole.Recruiter },

  // IT Admin role
  { groupName: 'IT Administrators', role: UserRole.ITAdmin },
  { groupName: 'IT Support', role: UserRole.ITAdmin },
  { groupName: 'JML IT Admins', role: UserRole.ITAdmin },

  // Executive role
  { groupName: 'Executives', role: UserRole.Executive },
  { groupName: 'Leadership Team', role: UserRole.Executive },
  { groupName: 'C-Suite', role: UserRole.Executive },

  // Site Admin role - highest privileges
  { groupName: 'Site Owners', role: UserRole.SiteAdmin },
  { groupName: 'Site Administrators', role: UserRole.SiteAdmin },
  { groupName: 'JML Admins', role: UserRole.SiteAdmin },

  // Procurement Manager role
  { groupName: 'Procurement Managers', role: UserRole.ProcurementManager },
  { groupName: 'Procurement Team', role: UserRole.ProcurementManager },
  { groupName: 'JML Procurement', role: UserRole.ProcurementManager },

  // Skills Manager role
  { groupName: 'Skills Managers', role: UserRole.SkillsManager },
  { groupName: 'Training Team', role: UserRole.SkillsManager },
  { groupName: 'Learning & Development', role: UserRole.SkillsManager },
  { groupName: 'JML Skills Managers', role: UserRole.SkillsManager },

  // Contract Manager role
  { groupName: 'Contract Managers', role: UserRole.ContractManager },
  { groupName: 'Legal Team', role: UserRole.ContractManager },
  { groupName: 'Contracts Team', role: UserRole.ContractManager },
  { groupName: 'JML Contract Managers', role: UserRole.ContractManager },

  // Finance Admin role
  { groupName: 'Finance Administrators', role: UserRole.FinanceAdmin },
  { groupName: 'Finance Team', role: UserRole.FinanceAdmin },
  { groupName: 'Finance Managers', role: UserRole.FinanceAdmin },
  { groupName: 'Accounting Team', role: UserRole.FinanceAdmin },
  { groupName: 'JML Finance Admins', role: UserRole.FinanceAdmin }
];

/**
 * Service for detecting user roles based on SharePoint group membership
 */
export class RoleDetectionService {
  private sp: SPFI;
  private roleMappings: IRoleMapping[];
  private userRolesCache: Map<string, UserRole[]> = new Map();
  private cacheExpiration: number = 5 * 60 * 1000; // 5 minutes

  constructor(sp: SPFI, customRoleMappings?: IRoleMapping[]) {
    this.sp = sp;
    this.roleMappings = customRoleMappings || DEFAULT_ROLE_MAPPINGS;
  }

  /**
   * Get all roles for the current user
   */
  public async getCurrentUserRoles(): Promise<UserRole[]> {
    try {
      const currentUser = await this.sp.web.currentUser();
      return await this.getUserRoles(currentUser.LoginName);
    } catch (error) {
      console.error('[RoleDetectionService] Error getting current user roles:', error);
      // Fallback to Employee role if detection fails
      return [UserRole.Employee];
    }
  }

  /**
   * Get all roles for a specific user by login name
   */
  public async getUserRoles(loginName: string): Promise<UserRole[]> {
    // Check cache first
    const cached = this.getCachedRoles(loginName);
    if (cached) {
      return cached;
    }

    try {
      // Get all groups the user belongs to
      const userGroups = await this.sp.web.siteUsers.getByLoginName(loginName).groups();

      // Map groups to roles
      const roles = new Set<UserRole>();

      for (const group of userGroups) {
        const mapping = this.roleMappings.find(m =>
          m.groupName.toLowerCase() === group.Title.toLowerCase()
        );

        if (mapping) {
          roles.add(mapping.role);
        }
      }

      // Check if user is site admin
      const user = await this.sp.web.siteUsers.getByLoginName(loginName)();
      if (user.IsSiteAdmin) {
        roles.add(UserRole.SiteAdmin);
      }

      // If no roles detected, default to Employee
      const userRoles = roles.size > 0 ? Array.from(roles) : [UserRole.Employee];

      // Cache the result
      this.cacheUserRoles(loginName, userRoles);

      return userRoles;
    } catch (error) {
      console.error('[RoleDetectionService] Error getting user roles:', error);
      // Fallback to Employee role
      return [UserRole.Employee];
    }
  }

  /**
   * Check if current user has a specific role
   */
  public async hasRole(role: UserRole): Promise<boolean> {
    const roles = await this.getCurrentUserRoles();
    return roles.includes(role);
  }

  /**
   * Check if current user has any of the specified roles
   */
  public async hasAnyRole(roles: UserRole[]): Promise<boolean> {
    const userRoles = await this.getCurrentUserRoles();
    return roles.some(role => userRoles.includes(role));
  }

  /**
   * Check if current user has all of the specified roles
   */
  public async hasAllRoles(roles: UserRole[]): Promise<boolean> {
    const userRoles = await this.getCurrentUserRoles();
    return roles.every(role => userRoles.includes(role));
  }

  /**
   * Get the highest priority role for the current user
   * Priority order: SiteAdmin > Executive > HRAdmin > ITAdmin > Recruiter > Manager > Employee
   */
  public async getPrimaryRole(): Promise<UserRole> {
    const roles = await this.getCurrentUserRoles();

    const priorityOrder = [
      UserRole.SiteAdmin,
      UserRole.Executive,
      UserRole.HRAdmin,
      UserRole.ITAdmin,
      UserRole.FinanceAdmin,
      UserRole.ProcurementManager,
      UserRole.ContractManager,
      UserRole.SkillsManager,
      UserRole.Recruiter,
      UserRole.Manager,
      UserRole.Employee
    ];

    for (const role of priorityOrder) {
      if (roles.includes(role)) {
        return role;
      }
    }

    return UserRole.Employee; // Fallback
  }

  /**
   * Update role mappings (useful for custom configurations)
   */
  public updateRoleMappings(mappings: IRoleMapping[]): void {
    this.roleMappings = mappings;
    this.clearCache();
  }

  /**
   * Clear the roles cache
   */
  public clearCache(): void {
    this.userRolesCache.clear();
  }

  /**
   * Get cached roles if available and not expired
   */
  private getCachedRoles(loginName: string): UserRole[] | null {
    const cached = this.userRolesCache.get(loginName);
    if (cached) {
      return cached;
    }
    return null;
  }

  /**
   * Cache user roles
   */
  private cacheUserRoles(loginName: string, roles: UserRole[]): void {
    this.userRolesCache.set(loginName, roles);

    // Set expiration timer
    setTimeout(() => {
      this.userRolesCache.delete(loginName);
    }, this.cacheExpiration);
  }

  /**
   * Get role display name
   */
  public static getRoleDisplayName(role: UserRole): string {
    const displayNames: Record<UserRole, string> = {
      [UserRole.Employee]: 'Employee',
      [UserRole.Manager]: 'Manager',
      [UserRole.HRAdmin]: 'HR Administrator',
      [UserRole.Recruiter]: 'Recruiter',
      [UserRole.ITAdmin]: 'IT Administrator',
      [UserRole.Executive]: 'Executive',
      [UserRole.SiteAdmin]: 'JML Administrator',
      [UserRole.ProcurementManager]: 'Procurement Manager',
      [UserRole.SkillsManager]: 'Skills Manager',
      [UserRole.ContractManager]: 'Contract Manager',
      [UserRole.FinanceAdmin]: 'Finance Administrator'
    };
    return displayNames[role];
  }

  /**
   * Get role description
   */
  public static getRoleDescription(role: UserRole): string {
    const descriptions: Record<UserRole, string> = {
      [UserRole.Employee]: 'Standard employee access to personal tasks and surveys',
      [UserRole.Manager]: 'Manage team members and approve JML processes',
      [UserRole.HRAdmin]: 'Full HR lifecycle management and reporting',
      [UserRole.Recruiter]: 'Talent acquisition and candidate management',
      [UserRole.ITAdmin]: 'IT asset management and system provisioning',
      [UserRole.Executive]: 'Executive dashboards and analytics',
      [UserRole.SiteAdmin]: 'Full system administration and configuration',
      [UserRole.ProcurementManager]: 'Procurement workflows, purchase orders, and vendor management',
      [UserRole.SkillsManager]: 'Training programs, certifications, and skills development',
      [UserRole.ContractManager]: 'Contract lifecycle, negotiations, and compliance tracking',
      [UserRole.FinanceAdmin]: 'Financial oversight, budgets, expenses, and payroll management'
    };
    return descriptions[role];
  }

  /**
   * Get role color for UI theming
   */
  public static getRoleColor(role: UserRole): string {
    const colors: Record<UserRole, string> = {
      [UserRole.Employee]: '#0078d4', // Blue
      [UserRole.Manager]: '#8764b8', // Purple
      [UserRole.HRAdmin]: '#00bcf2', // Cyan
      [UserRole.Recruiter]: '#00cc6a', // Green
      [UserRole.ITAdmin]: '#ff8c00', // Orange
      [UserRole.Executive]: '#d83b01', // Red-Orange
      [UserRole.SiteAdmin]: '#c239b3', // Magenta
      [UserRole.ProcurementManager]: '#038387', // Teal
      [UserRole.SkillsManager]: '#498205', // Olive Green
      [UserRole.ContractManager]: '#4f6bed', // Indigo
      [UserRole.FinanceAdmin]: '#004e8c'  // Dark Blue
    };
    return colors[role];
  }
}
