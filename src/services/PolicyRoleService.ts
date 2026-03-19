/**
 * Policy Manager Role Service
 * Maps organization-wide UserRole to Policy Manager-specific roles
 * and provides navigation filtering based on role hierarchy.
 */

import { UserRole } from './RoleDetectionService';

// ============================================================================
// POLICY MANAGER ROLES
// ============================================================================

/**
 * 4-tier role hierarchy for Policy Manager
 * User < Author < Manager < Admin
 */
export enum PolicyManagerRole {
  User = 'User',
  Author = 'Author',
  Manager = 'Manager',
  Admin = 'Admin'
}

// ============================================================================
// ROLE MAPPING
// ============================================================================

/**
 * Maps organization-wide UserRole to PolicyManagerRole
 */
const ROLE_MAPPING: Record<UserRole, PolicyManagerRole> = {
  [UserRole.Employee]: PolicyManagerRole.User,
  [UserRole.Recruiter]: PolicyManagerRole.Author,
  [UserRole.SkillsManager]: PolicyManagerRole.Author,
  [UserRole.ContractManager]: PolicyManagerRole.Author,
  [UserRole.Manager]: PolicyManagerRole.Manager,
  [UserRole.ProcurementManager]: PolicyManagerRole.Manager,
  [UserRole.HRAdmin]: PolicyManagerRole.Admin,
  [UserRole.ITAdmin]: PolicyManagerRole.Admin,
  [UserRole.FinanceAdmin]: PolicyManagerRole.Admin,
  [UserRole.Executive]: PolicyManagerRole.Admin,
  [UserRole.SiteAdmin]: PolicyManagerRole.Admin,
};

/**
 * Role hierarchy level (higher = more access)
 */
const ROLE_LEVEL: Record<PolicyManagerRole, number> = {
  [PolicyManagerRole.User]: 0,
  [PolicyManagerRole.Author]: 1,
  [PolicyManagerRole.Manager]: 2,
  [PolicyManagerRole.Admin]: 3,
};

// ============================================================================
// NAV KEY VISIBILITY
// ============================================================================

/**
 * Minimum role required to see each nav key
 */
// NAV_KEY_MIN_ROLE removed — permissions are now explicit per role via Admin > Role Permissions

/**
 * Header action visibility per role
 */
interface IHeaderVisibility {
  showSearch: boolean;
  showNotifications: boolean;
  showHelp: boolean;
  showSettings: boolean;
}

// Header visibility is now driven by explicit permissions (getHeaderVisibility function below)

// ============================================================================
// PUBLIC API
// ============================================================================

/**
 * Convert a UserRole (from RoleDetectionService) to a PolicyManagerRole
 */
export function toPolicyManagerRole(userRole: UserRole): PolicyManagerRole {
  return ROLE_MAPPING[userRole] || PolicyManagerRole.User;
}

/**
 * Get the highest PolicyManagerRole from an array of UserRoles
 */
export function getHighestPolicyRole(userRoles: UserRole[]): PolicyManagerRole {
  if (!userRoles || userRoles.length === 0) return PolicyManagerRole.User;

  let highest = PolicyManagerRole.User;
  for (const ur of userRoles) {
    const pmRole = toPolicyManagerRole(ur);
    if (ROLE_LEVEL[pmRole] > ROLE_LEVEL[highest]) {
      highest = pmRole;
    }
  }
  return highest;
}

/**
 * Check if a role has access to a given level
 */
export function hasMinimumRole(currentRole: PolicyManagerRole, requiredRole: PolicyManagerRole): boolean {
  return ROLE_LEVEL[currentRole] >= ROLE_LEVEL[requiredRole];
}


/**
 * Role permission entry as configured in Admin > Role Permissions
 */
export interface IRolePermissionEntry {
  feature: string;
  key: string;
  user: boolean;
  author: boolean;
  manager: boolean;
  admin: boolean;
  [roleKey: string]: string | boolean; // Index signature for dynamic role access
}

/**
 * Filter nav items based on HIERARCHICAL role access.
 * Higher roles inherit all access from lower roles:
 *   Admin >= Manager >= Author >= User
 *
 * Each nav item has a minimum role requirement. If the user's role
 * meets or exceeds it, the item is visible.
 *
 * @param navItems - The nav items to filter
 * @param role - The user's highest PolicyManagerRole
 * @param _permissions - Reserved for future use (admin-configurable overrides)
 */
export function filterNavForRole<T extends { key: string }>(
  navItems: T[],
  role: PolicyManagerRole,
  _permissions?: IRolePermissionEntry[] | null
): T[] {
  return navItems.filter(item => {
    const minRole = NAV_MINIMUM_ROLE[item.key];
    if (minRole === undefined) return role === PolicyManagerRole.Admin;
    return hasMinimumRole(role, minRole);
  });
}

/**
 * Minimum role required for each nav item key.
 * Higher roles automatically have access to everything below them.
 */
const NAV_MINIMUM_ROLE: Record<string, PolicyManagerRole> = {
  'browse':        PolicyManagerRole.User,
  'my-policies':   PolicyManagerRole.User,
  'details':       PolicyManagerRole.User,
  'search':        PolicyManagerRole.User,
  'help':          PolicyManagerRole.User,
  'create':        PolicyManagerRole.Author,
  'author':        PolicyManagerRole.Author,
  'packs':         PolicyManagerRole.Author,
  'quiz':          PolicyManagerRole.Author,
  'requests':      PolicyManagerRole.Author,
  'distribution':  PolicyManagerRole.Manager,
  'analytics':     PolicyManagerRole.Manager,
  'manager':       PolicyManagerRole.Manager,
  'approvals':     PolicyManagerRole.Manager,
  'delegations':   PolicyManagerRole.Manager,
  'reports':       PolicyManagerRole.Manager,
  'executive':     PolicyManagerRole.Manager,
  'admin':         PolicyManagerRole.Admin,
};

/**
 * Default permission table — used when admin hasn't configured custom permissions
 */
export function getDefaultPermissions(): IRolePermissionEntry[] {
  return [
    { feature: 'Browse Policies', key: 'browse', user: true, author: true, manager: true, admin: true },
    { feature: 'My Policies', key: 'myPolicies', user: true, author: true, manager: true, admin: true },
    { feature: 'Policy Details', key: 'details', user: true, author: true, manager: true, admin: true },
    { feature: 'Create Policy', key: 'create', user: false, author: true, manager: false, admin: true },
    { feature: 'Edit Policy', key: 'edit', user: false, author: true, manager: false, admin: true },
    { feature: 'Delete Policy', key: 'delete', user: false, author: false, manager: false, admin: true },
    { feature: 'Policy Packs', key: 'packs', user: false, author: true, manager: false, admin: true },
    { feature: 'Approvals', key: 'approvals', user: false, author: false, manager: true, admin: true },
    { feature: 'Delegations', key: 'delegations', user: false, author: false, manager: true, admin: true },
    { feature: 'Distribution', key: 'distribution', user: false, author: false, manager: true, admin: true },
    { feature: 'Analytics', key: 'analytics', user: false, author: false, manager: true, admin: true },
    { feature: 'Quiz Builder', key: 'quizBuilder', user: false, author: true, manager: false, admin: true },
    { feature: 'Admin Centre', key: 'adminPanel', user: false, author: false, manager: false, admin: true },
    { feature: 'User Management', key: 'userMgmt', user: false, author: false, manager: false, admin: true },
    { feature: 'System Settings', key: 'settings', user: false, author: false, manager: false, admin: true },
  ];
}

/**
 * Get header action visibility based on hierarchical role.
 * Settings cog visible for Manager and above.
 */
export function getHeaderVisibility(role: PolicyManagerRole, _permissions?: IRolePermissionEntry[] | null): IHeaderVisibility {
  return {
    showSearch: true,
    showNotifications: true,
    showHelp: true,
    showSettings: hasMinimumRole(role, PolicyManagerRole.Manager)
  };
}

/**
 * Get display label for a PolicyManagerRole
 */
export function getRoleDisplayName(role: PolicyManagerRole): string {
  switch (role) {
    case PolicyManagerRole.User: return 'User';
    case PolicyManagerRole.Author: return 'Policy Author';
    case PolicyManagerRole.Manager: return 'Manager';
    case PolicyManagerRole.Admin: return 'Administrator';
    default: return 'User';
  }
}

/**
 * Get role badge color (for header display)
 */
export function getRoleBadgeColor(role: PolicyManagerRole): string {
  switch (role) {
    case PolicyManagerRole.User: return '#64748b';
    case PolicyManagerRole.Author: return '#0d9488';
    case PolicyManagerRole.Manager: return '#f59e0b';
    case PolicyManagerRole.Admin: return '#ef4444';
    default: return '#64748b';
  }
}
