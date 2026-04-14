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
  // Policy Manager specific groups (PM_ prefix)
  { groupName: 'PM_PolicyAuthors', role: UserRole.Recruiter },     // Maps to Author via PolicyRoleService
  { groupName: 'PM_PolicyManagers', role: UserRole.Manager },
  { groupName: 'PM_PolicyAdmins', role: UserRole.SiteAdmin },

  // Standard SharePoint site groups (auto-created by SP)
  { groupName: 'PolicyManager Owners', role: UserRole.SiteAdmin },
  { groupName: 'PolicyManager Members', role: UserRole.Employee },
  { groupName: 'PolicyManager Visitors', role: UserRole.Employee },

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

  /** sessionStorage key for cross-webpart role caching */
  private static readonly SESSION_CACHE_KEY = 'pm_detected_roles';
  private static readonly SESSION_CACHE_TTL = 5 * 60 * 1000; // 5 minutes

  constructor(sp: SPFI, customRoleMappings?: IRoleMapping[]) {
    this.sp = sp;
    this.roleMappings = customRoleMappings || DEFAULT_ROLE_MAPPINGS;
  }

  /**
   * Get all roles for the current user.
   * Uses sessionStorage to cache across webpart instances on the same page.
   */
  public async getCurrentUserRoles(): Promise<UserRole[]> {
    try {
      // Check sessionStorage first (cross-webpart cache)
      const sessionCached = this.getSessionCachedRoles();
      if (sessionCached) {
        return sessionCached;
      }

      const currentUser = await this.sp.web.currentUser();
      const email = currentUser.Email || currentUser.LoginName || '';
      const roles = new Set<UserRole>();

      // ══════════════════════════════════════════════════════════════
      // ROLE DETECTION PRIORITY (highest wins):
      //
      //   1. PM_UserProfiles.PMRole — Admin Centre is the primary source
      //      Admin assigns role → stored in PMRole column → read here
      //
      //   2. Site Collection Admin — IsSiteAdmin = true → Admin role
      //      Can be overridden by PM_UserProfiles if a lower role is set
      //      (e.g. Site Admin who should only be Author in PolicyIQ)
      //
      //   3. Default — User role (no PM_UserProfiles entry = basic access)
      //
      // SP Security Groups (PM_PolicyAdmins, PM_PolicyAuthors, etc.)
      // are synced FOR REFERENCE ONLY when admin assigns roles. They
      // are NOT used for role detection — PM_UserProfiles is the source.
      // ══════════════════════════════════════════════════════════════

      let detectedRole = 'User'; // default

      // Step 1: Check PM_UserProfiles (Admin Centre assignment)
      try {
        if (email) {
          const profiles = await this.sp.web.lists.getByTitle('PM_UserProfiles')
            .items.filter(`EMail eq '${email.replace(/'/g, "''")}'`)
            .select('PMRole', 'PMRoles')
            .top(1)();
          if (profiles.length > 0) {
            const pmRole = (profiles[0].PMRole || '').trim();
            const pmRoles = (profiles[0].PMRoles || '').trim();
            const allRoleStrings = [pmRole, ...pmRoles.split(';')].map((r: string) => r.trim()).filter(Boolean);

            if (allRoleStrings.length > 0) {
              detectedRole = allRoleStrings[0]; // Primary role
            }

            // Store raw role strings for permission table lookups
            localStorage.setItem('pm_detected_role', detectedRole);
            localStorage.setItem('pm_detected_roles_all', allRoleStrings.join(';'));
          }
        }
      } catch {
        // PM_UserProfiles may not exist yet
      }

      // Step 2: Site Collection Admin gets Admin IF no PM_UserProfiles role is set
      // (If admin explicitly assigned a lower role in PM_UserProfiles, respect that)
      if (detectedRole === 'User' && currentUser.IsSiteAdmin) {
        detectedRole = 'Admin';
        localStorage.setItem('pm_detected_role', 'Admin');
      }

      // Map to UserRole enum
      const roleMap: Record<string, UserRole> = {
        'Admin': UserRole.SiteAdmin,
        'Manager': UserRole.Manager,
        'Author': UserRole.Recruiter,
        'User': UserRole.Employee
      };
      const mapped = roleMap[detectedRole];
      if (mapped) roles.add(mapped);

      const finalRoles = roles.size > 0 ? Array.from(roles) : [UserRole.Employee];

      // Cache in sessionStorage for other webparts
      this.setSessionCachedRoles(finalRoles);

      return finalRoles;
    } catch (error) {
      console.error('[RoleDetectionService] Error getting current user roles:', error);
      return [UserRole.Employee];
    }
  }

  /**
   * Read roles from sessionStorage if not expired
   */
  private getSessionCachedRoles(): UserRole[] | null {
    try {
      const raw = sessionStorage.getItem(RoleDetectionService.SESSION_CACHE_KEY);
      if (!raw) return null;
      const cached = JSON.parse(raw);
      if (cached && cached.expiry > Date.now() && Array.isArray(cached.roles)) {
        return cached.roles as UserRole[];
      }
      sessionStorage.removeItem(RoleDetectionService.SESSION_CACHE_KEY);
    } catch { /* ignore parse errors */ }
    return null;
  }

  /**
   * Write roles to sessionStorage with TTL
   */
  private setSessionCachedRoles(roles: UserRole[]): void {
    try {
      sessionStorage.setItem(
        RoleDetectionService.SESSION_CACHE_KEY,
        JSON.stringify({ roles, expiry: Date.now() + RoleDetectionService.SESSION_CACHE_TTL })
      );
    } catch { /* ignore quota errors */ }
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
    try { sessionStorage.removeItem(RoleDetectionService.SESSION_CACHE_KEY); } catch { /* ignore */ }
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
