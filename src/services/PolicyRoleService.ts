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
const NAV_KEY_MIN_ROLE: Record<string, PolicyManagerRole> = {
  'browse': PolicyManagerRole.User,
  'my-policies': PolicyManagerRole.User,
  'details': PolicyManagerRole.User,
  'create': PolicyManagerRole.Author,
  'packs': PolicyManagerRole.Author,
  'requests': PolicyManagerRole.Author,
  'author': PolicyManagerRole.Author,
  'manager': PolicyManagerRole.Manager,
  'approvals': PolicyManagerRole.Manager,
  'delegations': PolicyManagerRole.Manager,
  'distribution': PolicyManagerRole.Manager,
  'analytics': PolicyManagerRole.Manager,
  'quiz': PolicyManagerRole.Admin,
};

/**
 * Header action visibility per role
 */
interface IHeaderVisibility {
  showSearch: boolean;
  showNotifications: boolean;
  showHelp: boolean;
  showSettings: boolean;
}

const HEADER_VISIBILITY: Record<PolicyManagerRole, IHeaderVisibility> = {
  [PolicyManagerRole.User]: {
    showSearch: true,
    showNotifications: true,
    showHelp: true,
    showSettings: false,
  },
  [PolicyManagerRole.Author]: {
    showSearch: true,
    showNotifications: true,
    showHelp: true,
    showSettings: false,
  },
  [PolicyManagerRole.Manager]: {
    showSearch: true,
    showNotifications: true,
    showHelp: true,
    showSettings: true,
  },
  [PolicyManagerRole.Admin]: {
    showSearch: true,
    showNotifications: true,
    showHelp: true,
    showSettings: true,
  },
};

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
 * Filter nav items based on the user's PolicyManagerRole.
 * Returns only nav items the user is allowed to see.
 */
export function filterNavForRole<T extends { key: string }>(navItems: T[], role: PolicyManagerRole): T[] {
  return navItems.filter(item => {
    const minRole = NAV_KEY_MIN_ROLE[item.key];
    if (!minRole) return true; // Unknown keys default to visible
    return ROLE_LEVEL[role] >= ROLE_LEVEL[minRole];
  });
}

/**
 * Get header action visibility for a role
 */
export function getHeaderVisibility(role: PolicyManagerRole): IHeaderVisibility {
  return HEADER_VISIBILITY[role] || HEADER_VISIBILITY[PolicyManagerRole.User];
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
