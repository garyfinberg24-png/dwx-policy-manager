/**
 * PolicyRoleService Unit Tests
 *
 * Tests role hierarchy, mapToRole, role comparison, nav filtering,
 * header visibility, display names, and badge colors.
 */

// Mock RoleDetectionService to provide the UserRole enum
jest.mock('../RoleDetectionService', () => ({
  UserRole: {
    Employee: 'Employee',
    Manager: 'Manager',
    HRAdmin: 'HRAdmin',
    Recruiter: 'Recruiter',
    ITAdmin: 'ITAdmin',
    Executive: 'Executive',
    SiteAdmin: 'SiteAdmin',
    ProcurementManager: 'ProcurementManager',
    SkillsManager: 'SkillsManager',
    ContractManager: 'ContractManager',
    FinanceAdmin: 'FinanceAdmin',
  },
}));

import { UserRole } from '../RoleDetectionService';
import {
  PolicyManagerRole,
  toPolicyManagerRole,
  getHighestPolicyRole,
  hasMinimumRole,
  filterNavForRole,
  getHeaderVisibility,
  getRoleDisplayName,
  getRoleBadgeColor,
} from '../PolicyRoleService';

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('PolicyRoleService', () => {

  // ===== PolicyManagerRole enum =====

  describe('PolicyManagerRole enum', () => {
    it('should have 4 roles', () => {
      expect(PolicyManagerRole.User).toBe('User');
      expect(PolicyManagerRole.Author).toBe('Author');
      expect(PolicyManagerRole.Manager).toBe('Manager');
      expect(PolicyManagerRole.Admin).toBe('Admin');
    });
  });

  // ===== toPolicyManagerRole =====

  describe('toPolicyManagerRole', () => {
    it('should map Employee to User', () => {
      expect(toPolicyManagerRole(UserRole.Employee)).toBe(PolicyManagerRole.User);
    });

    it('should map Recruiter to Author', () => {
      expect(toPolicyManagerRole(UserRole.Recruiter)).toBe(PolicyManagerRole.Author);
    });

    it('should map SkillsManager to Author', () => {
      expect(toPolicyManagerRole(UserRole.SkillsManager)).toBe(PolicyManagerRole.Author);
    });

    it('should map ContractManager to Author', () => {
      expect(toPolicyManagerRole(UserRole.ContractManager)).toBe(PolicyManagerRole.Author);
    });

    it('should map Manager to Manager', () => {
      expect(toPolicyManagerRole(UserRole.Manager)).toBe(PolicyManagerRole.Manager);
    });

    it('should map ProcurementManager to Manager', () => {
      expect(toPolicyManagerRole(UserRole.ProcurementManager)).toBe(PolicyManagerRole.Manager);
    });

    it('should map HRAdmin to Admin', () => {
      expect(toPolicyManagerRole(UserRole.HRAdmin)).toBe(PolicyManagerRole.Admin);
    });

    it('should map ITAdmin to Admin', () => {
      expect(toPolicyManagerRole(UserRole.ITAdmin)).toBe(PolicyManagerRole.Admin);
    });

    it('should map Executive to Admin', () => {
      expect(toPolicyManagerRole(UserRole.Executive)).toBe(PolicyManagerRole.Admin);
    });

    it('should map SiteAdmin to Admin', () => {
      expect(toPolicyManagerRole(UserRole.SiteAdmin)).toBe(PolicyManagerRole.Admin);
    });

    it('should map FinanceAdmin to Admin', () => {
      expect(toPolicyManagerRole(UserRole.FinanceAdmin)).toBe(PolicyManagerRole.Admin);
    });

    it('should default to User for an unknown role', () => {
      expect(toPolicyManagerRole('UnknownRole' as UserRole)).toBe(PolicyManagerRole.User);
    });
  });

  // ===== getHighestPolicyRole =====

  describe('getHighestPolicyRole', () => {
    it('should return User for empty array', () => {
      expect(getHighestPolicyRole([])).toBe(PolicyManagerRole.User);
    });

    it('should return User for null/undefined input', () => {
      expect(getHighestPolicyRole(null as any)).toBe(PolicyManagerRole.User);
      expect(getHighestPolicyRole(undefined as any)).toBe(PolicyManagerRole.User);
    });

    it('should return the highest role among multiple UserRoles', () => {
      const roles = [UserRole.Employee, UserRole.Manager, UserRole.Recruiter];
      expect(getHighestPolicyRole(roles)).toBe(PolicyManagerRole.Manager);
    });

    it('should return Admin when an Admin-level role is present', () => {
      const roles = [UserRole.Employee, UserRole.HRAdmin];
      expect(getHighestPolicyRole(roles)).toBe(PolicyManagerRole.Admin);
    });

    it('should return Author when only Author-level roles exist', () => {
      const roles = [UserRole.Recruiter, UserRole.SkillsManager];
      expect(getHighestPolicyRole(roles)).toBe(PolicyManagerRole.Author);
    });

    it('should return User for a single Employee role', () => {
      expect(getHighestPolicyRole([UserRole.Employee])).toBe(PolicyManagerRole.User);
    });
  });

  // ===== hasMinimumRole =====

  describe('hasMinimumRole', () => {
    it('Admin should have minimum role User', () => {
      expect(hasMinimumRole(PolicyManagerRole.Admin, PolicyManagerRole.User)).toBe(true);
    });

    it('Admin should have minimum role Admin', () => {
      expect(hasMinimumRole(PolicyManagerRole.Admin, PolicyManagerRole.Admin)).toBe(true);
    });

    it('User should NOT have minimum role Author', () => {
      expect(hasMinimumRole(PolicyManagerRole.User, PolicyManagerRole.Author)).toBe(false);
    });

    it('Author should have minimum role Author', () => {
      expect(hasMinimumRole(PolicyManagerRole.Author, PolicyManagerRole.Author)).toBe(true);
    });

    it('Manager should have minimum role Author', () => {
      expect(hasMinimumRole(PolicyManagerRole.Manager, PolicyManagerRole.Author)).toBe(true);
    });

    it('Author should NOT have minimum role Manager', () => {
      expect(hasMinimumRole(PolicyManagerRole.Author, PolicyManagerRole.Manager)).toBe(false);
    });
  });

  // ===== filterNavForRole =====

  describe('filterNavForRole', () => {
    const allNavItems = [
      { key: 'browse', text: 'Browse' },
      { key: 'my-policies', text: 'My Policies' },
      { key: 'create', text: 'Create' },
      { key: 'packs', text: 'Packs' },
      { key: 'author', text: 'Author' },
      { key: 'manager', text: 'Manager' },
      { key: 'approvals', text: 'Approvals' },
      { key: 'distribution', text: 'Distribution' },
      { key: 'analytics', text: 'Analytics' },
      { key: 'quiz', text: 'Quiz Builder' },
    ];

    it('User should only see browse and my-policies', () => {
      const filtered = filterNavForRole(allNavItems, PolicyManagerRole.User);
      const keys = filtered.map(i => i.key);
      expect(keys).toContain('browse');
      expect(keys).toContain('my-policies');
      expect(keys).not.toContain('create');
      expect(keys).not.toContain('quiz');
      expect(keys).not.toContain('approvals');
    });

    it('Author should see User items plus create, packs, author', () => {
      const filtered = filterNavForRole(allNavItems, PolicyManagerRole.Author);
      const keys = filtered.map(i => i.key);
      expect(keys).toContain('browse');
      expect(keys).toContain('create');
      expect(keys).toContain('packs');
      expect(keys).toContain('author');
      expect(keys).not.toContain('approvals');
      expect(keys).not.toContain('quiz');
    });

    it('Manager should see Author items plus approvals, distribution, analytics, manager', () => {
      const filtered = filterNavForRole(allNavItems, PolicyManagerRole.Manager);
      const keys = filtered.map(i => i.key);
      expect(keys).toContain('browse');
      expect(keys).toContain('create');
      expect(keys).toContain('approvals');
      expect(keys).toContain('distribution');
      expect(keys).toContain('analytics');
      expect(keys).toContain('manager');
      expect(keys).not.toContain('quiz');
    });

    it('Admin should see all items including quiz', () => {
      const filtered = filterNavForRole(allNavItems, PolicyManagerRole.Admin);
      const keys = filtered.map(i => i.key);
      expect(keys).toHaveLength(allNavItems.length);
      expect(keys).toContain('quiz');
    });

    it('should default unknown keys to visible', () => {
      const items = [{ key: 'unknown-section', text: 'Unknown' }];
      const filtered = filterNavForRole(items, PolicyManagerRole.User);
      expect(filtered).toHaveLength(1);
    });

    it('should return empty array for empty input', () => {
      const filtered = filterNavForRole([], PolicyManagerRole.Admin);
      expect(filtered).toEqual([]);
    });
  });

  // ===== getHeaderVisibility =====

  describe('getHeaderVisibility', () => {
    it('User should NOT see settings', () => {
      const vis = getHeaderVisibility(PolicyManagerRole.User);
      expect(vis.showSearch).toBe(true);
      expect(vis.showNotifications).toBe(true);
      expect(vis.showHelp).toBe(true);
      expect(vis.showSettings).toBe(false);
    });

    it('Author should NOT see settings', () => {
      const vis = getHeaderVisibility(PolicyManagerRole.Author);
      expect(vis.showSettings).toBe(false);
    });

    it('Manager should see settings', () => {
      const vis = getHeaderVisibility(PolicyManagerRole.Manager);
      expect(vis.showSettings).toBe(true);
    });

    it('Admin should see settings', () => {
      const vis = getHeaderVisibility(PolicyManagerRole.Admin);
      expect(vis.showSettings).toBe(true);
    });
  });

  // ===== getRoleDisplayName =====

  describe('getRoleDisplayName', () => {
    it('should return correct display names', () => {
      expect(getRoleDisplayName(PolicyManagerRole.User)).toBe('User');
      expect(getRoleDisplayName(PolicyManagerRole.Author)).toBe('Policy Author');
      expect(getRoleDisplayName(PolicyManagerRole.Manager)).toBe('Manager');
      expect(getRoleDisplayName(PolicyManagerRole.Admin)).toBe('Administrator');
    });

    it('should default to User for unknown role', () => {
      expect(getRoleDisplayName('Unknown' as PolicyManagerRole)).toBe('User');
    });
  });

  // ===== getRoleBadgeColor =====

  describe('getRoleBadgeColor', () => {
    it('should return hex color codes for all roles', () => {
      expect(getRoleBadgeColor(PolicyManagerRole.User)).toBe('#64748b');
      expect(getRoleBadgeColor(PolicyManagerRole.Author)).toBe('#0d9488');
      expect(getRoleBadgeColor(PolicyManagerRole.Manager)).toBe('#f59e0b');
      expect(getRoleBadgeColor(PolicyManagerRole.Admin)).toBe('#ef4444');
    });

    it('should default to User color for unknown role', () => {
      expect(getRoleBadgeColor('Unknown' as PolicyManagerRole)).toBe('#64748b');
    });
  });
});
