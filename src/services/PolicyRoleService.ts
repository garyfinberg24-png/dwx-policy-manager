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
 * Map from nav item keys to role permission table keys.
 * The permission table uses different keys than the nav items.
 */
const NAV_KEY_TO_PERMISSION_KEY: Record<string, string> = {
  'browse': 'browse',
  'my-policies': 'myPolicies',
  'details': 'details',
  'create': 'create',
  'author': 'create',        // "Policy Author" page uses create permission
  'packs': 'packs',
  'requests': 'create',      // "Requests" falls under create permission
  'approvals': 'approvals',
  'delegations': 'delegations',
  'distribution': 'distribution',
  'manager': 'approvals',    // "Policy Manager" page uses approvals permission
  'analytics': 'analytics',
  'quiz': 'quizBuilder',
};

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
 * Filter nav items based on EXPLICIT role permissions.
 * No hierarchy inheritance — each role sees ONLY what's explicitly enabled for it.
 *
 * @param navItems - The nav items to filter
 * @param role - The user's role (or roles)
 * @param permissions - The admin-configured permission table (if null, falls back to defaults)
 */
export function filterNavForRole<T extends { key: string }>(
  navItems: T[],
  role: PolicyManagerRole,
  permissions?: IRolePermissionEntry[] | null
): T[] {
  // If no explicit permissions configured, use the default permission table
  const permTable = permissions && permissions.length > 0 ? permissions : getDefaultPermissions();
  const roleKey = role.toLowerCase(); // 'user', 'author', 'manager', 'admin'

  return navItems.filter(item => {
    const permKey = NAV_KEY_TO_PERMISSION_KEY[item.key];
    if (!permKey) return true; // Unknown nav keys default to visible

    // Find the permission entry for this feature
    const entry = permTable.find(p => p.key === permKey);
    if (!entry) return true; // No permission entry = visible by default

    // Admin always has access
    if (role === PolicyManagerRole.Admin) return true;

    // Check explicit permission for this role — NO inheritance
    return entry[roleKey] === true;
  });
}

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
 * Get header action visibility based on explicit permissions.
 */
export function getHeaderVisibility(role: PolicyManagerRole, permissions?: IRolePermissionEntry[] | null): IHeaderVisibility {
  const permTable = permissions && permissions.length > 0 ? permissions : getDefaultPermissions();
  const roleKey = role.toLowerCase();

  // Settings cog visible if 'adminPanel' or 'settings' permission is enabled for this role
  const settingsEntry = permTable.find(p => p.key === 'adminPanel' || p.key === 'settings');
  const showSettings = role === PolicyManagerRole.Admin || (settingsEntry ? settingsEntry[roleKey] === true : false);

  return {
    showSearch: true,
    showNotifications: true,
    showHelp: true,
    showSettings
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
