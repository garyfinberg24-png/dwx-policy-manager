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
 * Filter nav items based on EXPLICIT permission table with hierarchy fallback.
 * The explicit permission table (from getDefaultPermissions()) takes precedence.
 * For nav keys not found in the permission table, falls back to the
 * NAV_MINIMUM_ROLE hierarchy check.
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
  // Build a lookup from the explicit permission table
  const permEntries = getDefaultPermissions();
  const permMap = new Map<string, IRolePermissionEntry>();
  for (const entry of permEntries) {
    permMap.set(entry.key, entry);
  }

  const roleKey = role.toLowerCase(); // 'user' | 'author' | 'manager' | 'admin'

  return navItems.filter(item => {
    // Check explicit permission table first
    const perm = permMap.get(item.key);
    if (perm) {
      return perm[roleKey] === true;
    }

    // Fallback to hierarchy check for keys not in the permission table
    const minRole = NAV_MINIMUM_ROLE[item.key];
    if (minRole === undefined) return role === PolicyManagerRole.Admin;
    return hasMinimumRole(role, minRole);
  });
}

/**
 * Minimum role required for each nav item key.
 * Used as FALLBACK when a key is not in the explicit permission table.
 * Keys not listed here default to Admin-only.
 */
const NAV_MINIMUM_ROLE: Record<string, PolicyManagerRole> = {
  // ── User ──
  'browse':           PolicyManagerRole.User,
  'my-policies':      PolicyManagerRole.User,
  'details':          PolicyManagerRole.User,
  'search':           PolicyManagerRole.User,
  'help':             PolicyManagerRole.User,
  // ── Author ──
  'newpolicy':        PolicyManagerRole.Author,
  'create':           PolicyManagerRole.Author,
  'author':           PolicyManagerRole.Author,
  'packs':            PolicyManagerRole.Author,
  'quiz':             PolicyManagerRole.Author,
  'author-reports':   PolicyManagerRole.Author,
  'bulk-upload':      PolicyManagerRole.Author,
  'requests':         PolicyManagerRole.Author,
  // ── Manager ──
  'manager-dashboard': PolicyManagerRole.Manager,
  'approvals':        PolicyManagerRole.Manager,
  'distribution':     PolicyManagerRole.Manager,
  'team-compliance':  PolicyManagerRole.Manager,
  'delegations':      PolicyManagerRole.Manager,
  'reviews':          PolicyManagerRole.Manager,
  'reports':          PolicyManagerRole.Manager,
  'analytics':        PolicyManagerRole.Manager,
  'request-policy':   PolicyManagerRole.Manager,
  'executive':        PolicyManagerRole.Manager,
  'manager':          PolicyManagerRole.Manager,
  // ── Admin ──
  'admin':            PolicyManagerRole.Admin,
  'eventviewer':      PolicyManagerRole.Admin,
};

/**
 * Default permission table — explicit per-role visibility for each nav item.
 * KEYS MUST MATCH the nav item keys in PolicyManagerHeader.tsx exactly.
 *
 * Role visibility rules:
 *   User:    My Policies, Policy Hub, Help, Search
 *   Author:  + New Policy, Drafts & Pipeline, Policy Packs, Quiz Builder, Reports (Author), Bulk Upload
 *            Author does NOT see: Manager items (Dashboard, Approvals, Distribution, Analytics, etc.)
 *   Manager: + Dashboard, Approvals, Distribution, Team Compliance, Delegations, Review Cycles,
 *            Reports (Manager), Analytics, Request Policy
 *            Manager does NOT see: Author items (New Policy, Drafts, Packs, Quiz, Author Reports, Bulk Upload)
 *   Admin:   Everything
 */
export function getDefaultPermissions(): IRolePermissionEntry[] {
  return [
    // ── Visible to ALL roles ──
    { feature: 'Policy Hub',        key: 'browse',           user: true,  author: true,  manager: true,  admin: true },
    { feature: 'My Policies',       key: 'my-policies',      user: true,  author: true,  manager: true,  admin: true },
    { feature: 'Policy Details',    key: 'details',          user: true,  author: true,  manager: true,  admin: true },
    { feature: 'Search',            key: 'search',           user: true,  author: true,  manager: true,  admin: true },
    { feature: 'Help',              key: 'help',             user: true,  author: true,  manager: true,  admin: true },
    // ── Author + Admin only ──
    { feature: 'New Policy',        key: 'newpolicy',        user: false, author: true,  manager: false, admin: true },
    { feature: 'Create Policy',     key: 'create',           user: false, author: true,  manager: false, admin: true },
    { feature: 'Drafts & Pipeline', key: 'author',           user: false, author: true,  manager: false, admin: true },
    { feature: 'Policy Packs',      key: 'packs',            user: false, author: true,  manager: false, admin: true },
    { feature: 'Quiz Builder',      key: 'quiz',             user: false, author: true,  manager: false, admin: true },
    { feature: 'Reports (Author)',  key: 'author-reports',   user: false, author: true,  manager: false, admin: true },
    { feature: 'Bulk Upload',       key: 'bulk-upload',      user: false, author: true,  manager: false, admin: true },
    // ── Manager + Admin only ──
    { feature: 'Dashboard',         key: 'manager-dashboard',user: false, author: false, manager: true,  admin: true },
    { feature: 'Approvals',         key: 'approvals',        user: false, author: false, manager: true,  admin: true },
    { feature: 'Distribution',      key: 'distribution',     user: false, author: false, manager: true,  admin: true },
    { feature: 'Team Compliance',   key: 'team-compliance',  user: false, author: false, manager: true,  admin: true },
    { feature: 'Delegations',       key: 'delegations',      user: false, author: false, manager: true,  admin: true },
    { feature: 'Review Cycles',     key: 'reviews',          user: false, author: false, manager: true,  admin: true },
    { feature: 'Reports (Manager)', key: 'reports',          user: false, author: false, manager: true,  admin: true },
    { feature: 'Analytics',         key: 'analytics',        user: false, author: false, manager: true,  admin: true },
    { feature: 'Request Policy',    key: 'request-policy',   user: false, author: false, manager: true,  admin: true },
    // ── Admin only ──
    { feature: 'Admin Centre',      key: 'admin',            user: false, author: false, manager: false, admin: true },
    { feature: 'Event Viewer',      key: 'eventviewer',      user: false, author: false, manager: false, admin: true },
    { feature: 'Edit Policy',       key: 'edit',             user: false, author: true,  manager: false, admin: true },
    { feature: 'Delete Policy',     key: 'delete',           user: false, author: false, manager: false, admin: true },
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
    showSettings: hasMinimumRole(role, PolicyManagerRole.Admin)
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
