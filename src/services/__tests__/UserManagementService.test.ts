/**
 * UserManagementService Unit Tests
 *
 * Tests constructor, employee CRUD, role summary, search, and SP group
 * management by mocking the PnP SP fluent API chain.
 */

import { UserManagementService } from '../UserManagementService';

// ---------------------------------------------------------------------------
// PnP SP mock builder
// ---------------------------------------------------------------------------

function createMockSp(resolvedValue: unknown = [], addResult?: unknown) {
  const itemsMock = {
    select: jest.fn().mockReturnThis(),
    filter: jest.fn().mockReturnThis(),
    orderBy: jest.fn().mockReturnThis(),
    skip: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolvedValue)),
    add: jest.fn().mockResolvedValue(addResult ?? { data: resolvedValue }),
    getById: jest.fn().mockReturnValue({
      select: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolvedValue)),
      update: jest.fn().mockResolvedValue(undefined),
      delete: jest.fn().mockResolvedValue(undefined),
    }),
  };

  const usersMock = {
    select: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolvedValue)),
    add: jest.fn().mockResolvedValue(undefined),
    removeById: jest.fn().mockResolvedValue(undefined),
  };

  const siteGroupsMock = {
    select: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolvedValue)),
    getById: jest.fn().mockReturnValue({ users: usersMock }),
    add: jest.fn().mockResolvedValue(addResult ?? { data: { Id: 1, Title: '', Description: '' } }),
  };

  const sp = {
    web: {
      lists: {
        getByTitle: jest.fn().mockReturnValue({ items: itemsMock }),
      },
      siteGroups: siteGroupsMock,
    },
  };

  return { sp, itemsMock, siteGroupsMock, usersMock };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('UserManagementService', () => {
  let service: UserManagementService;
  let mockSp: ReturnType<typeof createMockSp>;

  beforeEach(() => {
    jest.clearAllMocks();
    mockSp = createMockSp();
    service = new UserManagementService(mockSp.sp as any);
  });

  // ===== Constructor =====

  describe('constructor', () => {
    it('should create an instance without throwing', () => {
      expect(service).toBeDefined();
    });
  });

  // ===== getEmployees =====

  describe('getEmployees', () => {
    it('should call the PM_Employees list', async () => {
      await service.getEmployees();
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_Employees');
    });

    it('should return items and total count', async () => {
      const employees = [{ Id: 1, Title: 'Alice' }, { Id: 2, Title: 'Bob' }];
      const { sp } = createMockSp(employees);
      service = new UserManagementService(sp as any);

      const result = await service.getEmployees(1, 25);
      expect(result.items).toEqual(employees);
      expect(result.total).toBe(2);
    });

    it('should return empty result on error', async () => {
      const failSp = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                select: jest.fn().mockReturnThis(),
                filter: jest.fn().mockReturnThis(),
                orderBy: jest.fn().mockReturnThis(),
                skip: jest.fn().mockReturnThis(),
                top: jest.fn().mockReturnValue(jest.fn().mockRejectedValue(new Error('SP fail'))),
              },
            }),
          },
        },
      };
      service = new UserManagementService(failSp as any);
      const result = await service.getEmployees();
      expect(result).toEqual({ items: [], total: 0 });
    });
  });

  // ===== searchEmployees =====

  describe('searchEmployees', () => {
    it('should filter by substringof on Title, Email, Department', async () => {
      await service.searchEmployees('test');
      expect(mockSp.itemsMock.filter).toHaveBeenCalledWith(
        expect.stringContaining("substringof('test',Title)")
      );
    });

    it('should return matching employees', async () => {
      const matches = [{ Id: 5, Title: 'Test User' }];
      const { sp } = createMockSp(matches);
      service = new UserManagementService(sp as any);

      const result = await service.searchEmployees('Test');
      expect(result).toEqual(matches);
    });

    it('should return empty array on error', async () => {
      const failSp = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                select: jest.fn().mockReturnThis(),
                filter: jest.fn().mockReturnThis(),
                orderBy: jest.fn().mockReturnThis(),
                top: jest.fn().mockReturnValue(jest.fn().mockRejectedValue(new Error('fail'))),
              },
            }),
          },
        },
      };
      service = new UserManagementService(failSp as any);
      const result = await service.searchEmployees('x');
      expect(result).toEqual([]);
    });
  });

  // ===== updateUserRole =====

  describe('updateUserRole', () => {
    it('should update PMRole on the correct item', async () => {
      await service.updateUserRole(10, 'Admin');
      expect(mockSp.itemsMock.getById).toHaveBeenCalledWith(10);
      expect(mockSp.itemsMock.getById(10).update).toHaveBeenCalledWith({ PMRole: 'Admin' });
    });

    it('should propagate errors', async () => {
      mockSp.itemsMock.getById.mockReturnValue({
        update: jest.fn().mockRejectedValue(new Error('update fail')),
      });
      await expect(service.updateUserRole(1, 'User')).rejects.toThrow('update fail');
    });
  });

  // ===== getRoleSummary =====

  describe('getRoleSummary', () => {
    it('should return counts per role', async () => {
      const items = [
        { PMRole: 'Admin' },
        { PMRole: 'Admin' },
        { PMRole: 'Author' },
        { PMRole: null },  // defaults to User
      ];
      const { sp } = createMockSp(items);
      service = new UserManagementService(sp as any);

      const result = await service.getRoleSummary();
      expect(result.find(r => r.role === 'Admin')?.count).toBe(2);
      expect(result.find(r => r.role === 'Author')?.count).toBe(1);
      expect(result.find(r => r.role === 'User')?.count).toBe(1);
      expect(result.find(r => r.role === 'Manager')?.count).toBe(0);
    });

    it('should return zero counts on error', async () => {
      const failSp = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                select: jest.fn().mockReturnThis(),
                top: jest.fn().mockReturnValue(jest.fn().mockRejectedValue(new Error('fail'))),
              },
            }),
          },
        },
      };
      service = new UserManagementService(failSp as any);
      const result = await service.getRoleSummary();
      expect(result).toHaveLength(4);
      expect(result.every(r => r.count === 0)).toBe(true);
    });
  });

  // ===== getSiteGroups =====

  describe('getSiteGroups', () => {
    it('should return all groups when no filter prefix given', async () => {
      const groups = [
        { Id: 1, Title: 'PM_Admins', Description: 'Admin group', OwnerTitle: 'Owner' },
        { Id: 2, Title: 'Other Group', Description: '', OwnerTitle: '' },
      ];
      const { sp } = createMockSp(groups);
      service = new UserManagementService(sp as any);

      const result = await service.getSiteGroups();
      expect(result).toHaveLength(2);
    });

    it('should filter groups by prefix', async () => {
      const groups = [
        { Id: 1, Title: 'PM_Admins', Description: '', OwnerTitle: '' },
        { Id: 2, Title: 'Other', Description: '', OwnerTitle: '' },
      ];
      const { sp } = createMockSp(groups);
      service = new UserManagementService(sp as any);

      const result = await service.getSiteGroups('PM_');
      expect(result).toHaveLength(1);
      expect(result[0].Title).toBe('PM_Admins');
    });
  });

  // ===== getGroupMembers =====

  describe('getGroupMembers', () => {
    it('should query the correct group by ID', async () => {
      const members = [
        { Id: 1, Title: 'Alice', Email: 'alice@test.com', LoginName: 'i:0#.f|alice', IsSiteAdmin: false },
      ];
      const { sp, siteGroupsMock, usersMock } = createMockSp(members);
      usersMock.select.mockReturnValue(jest.fn().mockResolvedValue(members));
      service = new UserManagementService(sp as any);

      const result = await service.getGroupMembers(5);
      expect(siteGroupsMock.getById).toHaveBeenCalledWith(5);
      expect(result).toHaveLength(1);
      expect(result[0].Email).toBe('alice@test.com');
    });

    it('should return empty array on error', async () => {
      mockSp.usersMock.select.mockReturnValue(jest.fn().mockRejectedValue(new Error('fail')));
      const result = await service.getGroupMembers(99);
      expect(result).toEqual([]);
    });
  });

  // ===== createGroup =====

  describe('createGroup', () => {
    it('should create a group and return the result', async () => {
      const addResult = { data: { Id: 42, Title: 'New Group', Description: 'Desc' } };
      const { sp } = createMockSp([], addResult);
      service = new UserManagementService(sp as any);

      const result = await service.createGroup('New Group', 'Desc');
      expect(result.Id).toBe(42);
      expect(result.Title).toBe('New Group');
    });
  });
});
