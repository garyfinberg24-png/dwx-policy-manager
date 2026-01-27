// @ts-nocheck
/**
 * PolicyService Unit Tests
 * Testing enterprise policy management service with mocked SharePoint dependencies
 */
/// <reference types="jest" />

// ============================================================================
// MOCKS - Must be defined BEFORE imports
// ============================================================================

// Mock @pnp/sp modules before they're imported
jest.mock('@pnp/sp', () => ({
  SPFI: jest.fn()
}));

jest.mock('@pnp/sp/webs', () => ({}));
jest.mock('@pnp/sp/lists', () => ({}));
jest.mock('@pnp/sp/items', () => ({}));
jest.mock('@pnp/sp/items/get-all', () => ({}));
jest.mock('@pnp/sp/site-users/web', () => ({}));
jest.mock('@pnp/sp/files', () => ({}));
jest.mock('@pnp/sp/folders', () => ({}));

// Now import modules (after mocks are defined)
import { PolicyService } from './PolicyService';
import {
  IPolicy,
  PolicyStatus,
  PolicyCategory,
  PolicyType,
  ComplianceRisk,
  VersionType,
  DocumentFormat,
  AcknowledgementType,
  DistributionScope,
  DataClassification,
  RetentionCategory,
  AcknowledgementStatus
} from '../models/IPolicy';
import { PolicyRole, PolicyOperation } from './PolicyValidationService';

// Mock LoggingService
jest.mock('./LoggingService', () => ({
  logger: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

// Mock PolicyCacheService - use singleton inside factory
jest.mock('./PolicyCacheService', () => {
  const actual = jest.requireActual('./PolicyCacheService');
  // Singleton instance - same object returned every time
  const singleton = {
    getPolicy: jest.fn(),
    setPolicy: jest.fn(),
    invalidatePolicy: jest.fn(),
    getPolicyList: jest.fn(),
    setPolicyList: jest.fn(),
    getStats: jest.fn().mockReturnValue({ hits: 0, misses: 0, size: 0, hitRate: 0 }),
    clear: jest.fn()
  };
  return {
    getPolicyCacheService: jest.fn(() => singleton),
    paginateArray: actual.paginateArray,
    generateCacheKey: actual.generateCacheKey
  };
});

// Get reference to the singleton for test access - use require to avoid TS type checking
const { getPolicyCacheService: getMockCacheService } = jest.requireMock('./PolicyCacheService');
const mockCacheService = getMockCacheService() as {
  getPolicy: jest.Mock;
  setPolicy: jest.Mock;
  invalidatePolicy: jest.Mock;
  getPolicyList: jest.Mock;
  setPolicyList: jest.Mock;
  getStats: jest.Mock;
  clear: jest.Mock;
};

// Shared mock instance for ValidationService - use var for hoisting (accessed in jest.mock factory)
// eslint-disable-next-line no-var
var mockValidationServiceInstance = {
  canPerformOperation: jest.fn().mockResolvedValue({ isValid: true }),
  validatePolicyData: jest.fn().mockResolvedValue({ isValid: true }),
  validateAcknowledgement: jest.fn().mockResolvedValue({ isValid: true })
};

// Mock ValidationService - return shared instance
jest.mock('./PolicyValidationService', () => {
  const MockValidationService = function() {
    return mockValidationServiceInstance;
  };

  return {
    PolicyValidationService: MockValidationService,
    PolicyOperation: {
      Create: 'Create',
      Read: 'Read',
      Update: 'Update',
      Delete: 'Delete',
      Approve: 'Approve',
      Reject: 'Reject',
      Publish: 'Publish',
      Archive: 'Archive',
      Acknowledge: 'Acknowledge'
    },
    PolicyRole: {
      Administrator: 'Administrator',
      ComplianceOfficer: 'ComplianceOfficer',
      Publisher: 'Publisher',
      Approver: 'Approver',
      Reviewer: 'Reviewer',
      Author: 'Author',
      Employee: 'Employee'
    }
  };
});

// Mock AuditService - define class inline to avoid hoisting issues
jest.mock('./PolicyAuditService', () => {
  // Define mock class inside factory function
  const MockAuditService = function() {
    this.initialize = jest.fn().mockResolvedValue(undefined);
    this.logEvent = jest.fn().mockResolvedValue(undefined);
    this.logPolicyApproval = jest.fn().mockResolvedValue(undefined);
    this.logPolicyRejection = jest.fn().mockResolvedValue(undefined);
    this.logAcknowledgement = jest.fn().mockResolvedValue(undefined);
  };

  return {
    PolicyAuditService: MockAuditService,
    AuditEventType: {
      PolicyCreated: 'PolicyCreated',
      PolicyUpdated: 'PolicyUpdated',
      PolicyPublished: 'PolicyPublished',
      PolicyApproved: 'PolicyApproved',
      PolicyRejected: 'PolicyRejected',
      UnauthorizedAccess: 'UnauthorizedAccess'
    },
    AuditSeverity: {
      Info: 'Info',
      Warning: 'Warning',
      Security: 'Security',
      Critical: 'Critical'
    }
  };
});

// Mock RetentionService - define inline to avoid hoisting issues
jest.mock('./PolicyRetentionService', () => {
  const MockRetentionService = function() {
    this.getRetentionPeriodDays = jest.fn().mockReturnValue(1095); // 3 years default
  };

  return {
    PolicyRetentionService: MockRetentionService
  };
});

// Create mock SP functions
const createMockSP = () => {
  const mockItemsArray: any[] = [];
  const mockAddResult = { data: { Id: 1 } };

  // Create a callable that also has chainable methods
  const createChainable = (resolveValue: any) => {
    const chainable: any = jest.fn().mockResolvedValue(resolveValue);
    chainable.select = jest.fn().mockReturnValue(chainable);
    chainable.expand = jest.fn().mockReturnValue(chainable);
    chainable.filter = jest.fn().mockReturnValue(chainable);
    chainable.orderBy = jest.fn().mockReturnValue(chainable);
    chainable.top = jest.fn().mockReturnValue(chainable);
    chainable.update = jest.fn().mockResolvedValue(undefined);
    return chainable;
  };

  const mockItems = {
    add: jest.fn().mockResolvedValue(mockAddResult),
    getById: jest.fn().mockImplementation((id: number) => {
      return createChainable({ Id: id, Title: 'Test Policy' });
    }),
    select: jest.fn().mockReturnThis(),
    expand: jest.fn().mockReturnThis(),
    filter: jest.fn().mockReturnThis(),
    orderBy: jest.fn().mockReturnThis(),
    top: jest.fn().mockResolvedValue(mockItemsArray)
  };
  // Make items callable
  (mockItems as any).__proto__ = jest.fn().mockResolvedValue(mockItemsArray);

  const mockList = {
    items: mockItems
  };

  const mockLists = {
    getByTitle: jest.fn().mockReturnValue(mockList)
  };

  // Create callable currentUser that also has methods
  const currentUserData = {
    Id: 1,
    Email: 'test@company.com',
    Title: 'Test User'
  };
  const mockCurrentUser: any = jest.fn().mockResolvedValue(currentUserData);
  mockCurrentUser.groups = jest.fn().mockResolvedValue([{ Title: 'Test Group' }]);

  const mockSiteUsers = jest.fn().mockResolvedValue([
    { Id: 1, Title: 'User 1' },
    { Id: 2, Title: 'User 2' }
  ]);

  const mockWeb = {
    lists: mockLists,
    currentUser: mockCurrentUser,
    siteUsers: mockSiteUsers
  };

  return {
    web: mockWeb,
    _mockItems: mockItemsArray,
    _mockAddResult: mockAddResult,
    _mockList: mockList,
    _mockCurrentUser: mockCurrentUser,
    _createChainable: createChainable
  };
};

// ============================================================================
// TEST DATA
// ============================================================================

const createMockPolicy = (overrides: Partial<IPolicy> = {}): IPolicy => ({
  Id: 1,
  Title: 'Test Policy',
  PolicyNumber: 'POL-HR-001',
  PolicyName: 'Test Policy',
  PolicyCategory: PolicyCategory.HRPolicies,
  PolicyType: PolicyType.Corporate,
  Description: 'Test policy description',
  VersionNumber: '1.0',
  VersionType: VersionType.Major,
  MajorVersion: 1,
  MinorVersion: 0,
  DocumentFormat: DocumentFormat.HTML,
  PolicyOwnerId: 1,
  PolicyAuthorIds: [1],
  Status: PolicyStatus.Draft,
  IsActive: false,
  IsMandatory: true,
  ComplianceRisk: ComplianceRisk.Medium,
  RequiresAcknowledgement: true,
  AcknowledgementType: AcknowledgementType.OneTime,
  RequiresQuiz: false,
  AllowRetake: false,
  DistributionScope: DistributionScope.AllEmployees,
  DataClassification: DataClassification.Internal,
  RetentionCategory: RetentionCategory.Standard,
  Created: new Date('2024-01-01'),
  Modified: new Date('2024-01-01'),
  ...overrides
});

// ============================================================================
// TESTS
// ============================================================================

describe('PolicyService', () => {
  let service: PolicyService;
  let mockSP: ReturnType<typeof createMockSP>;

  beforeEach(() => {
    // Clear mock call history but preserve implementations
    jest.clearAllMocks();

    // Re-establish mock implementations for cache service
    mockCacheService.getPolicy.mockReturnValue(null);
    mockCacheService.getPolicyList.mockReturnValue(null);
    mockCacheService.getStats.mockReturnValue({ hits: 0, misses: 0, size: 0, hitRate: 0 });

    // Re-establish mock implementations for validation service (shared instance)
    mockValidationServiceInstance.canPerformOperation.mockResolvedValue({ isValid: true });
    mockValidationServiceInstance.validatePolicyData.mockResolvedValue({ isValid: true });
    mockValidationServiceInstance.validateAcknowledgement.mockResolvedValue({ isValid: true });

    mockSP = createMockSP();
    service = new PolicyService(mockSP as any);
  });

  // ==========================================
  // Initialization Tests
  // ==========================================

  describe('initialize', () => {
    it('should initialize service with current user', async () => {
      await service.initialize();

      expect(mockSP.web.currentUser).toHaveBeenCalled();
      // Note: auditService.initialize is called internally - tested via successful initialization
    });

    it('should throw error if initialization fails', async () => {
      mockSP.web.currentUser.mockRejectedValueOnce(new Error('Auth failed'));

      await expect(service.initialize()).rejects.toThrow('Auth failed');
    });
  });

  // ==========================================
  // getCurrentUserRole Tests
  // ==========================================

  describe('getCurrentUserRole', () => {
    it('should return Administrator role for admin group', async () => {
      mockSP.web.currentUser.groups.mockResolvedValueOnce([{ Title: 'Site Administrators' }]);

      const role = await service.getCurrentUserRole();

      expect(role).toBe(PolicyRole.Administrator);
    });

    it('should return ComplianceOfficer role for compliance group', async () => {
      mockSP.web.currentUser.groups.mockResolvedValueOnce([{ Title: 'Compliance Team' }]);

      const role = await service.getCurrentUserRole();

      expect(role).toBe(PolicyRole.ComplianceOfficer);
    });

    it('should return Employee role by default', async () => {
      mockSP.web.currentUser.groups.mockResolvedValueOnce([{ Title: 'Regular Users' }]);

      const role = await service.getCurrentUserRole();

      expect(role).toBe(PolicyRole.Employee);
    });

    it('should return Employee role on error', async () => {
      mockSP.web.currentUser.groups.mockRejectedValueOnce(new Error('Failed'));

      const role = await service.getCurrentUserRole();

      expect(role).toBe(PolicyRole.Employee);
    });
  });

  // ==========================================
  // createPolicy Tests
  // ==========================================

  describe('createPolicy', () => {
    it('should create policy successfully', async () => {
      await service.initialize();

      const policyData: Partial<IPolicy> = {
        PolicyName: 'New Policy',
        PolicyCategory: PolicyCategory.HRPolicies,
        PolicyType: PolicyType.Corporate,
        DocumentFormat: DocumentFormat.HTML,
        AcknowledgementType: AcknowledgementType.OneTime,
        DistributionScope: DistributionScope.AllEmployees,
        ComplianceRisk: ComplianceRisk.Medium
      };

      // Mock getPolicyById to return the created policy
      const mockCreatedPolicy = createMockPolicy({ PolicyName: 'New Policy' });
      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(mockCreatedPolicy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      const result = await service.createPolicy(policyData);

      expect(mockSP.web.lists.getByTitle).toHaveBeenCalledWith('JML_Policies');
      expect(mockSP._mockList.items.add).toHaveBeenCalled();
      // Note: auditService.logEvent is called internally - tested via successful creation
    });

    it('should throw error if user lacks permission', async () => {
      await service.initialize();
      mockValidationServiceInstance.canPerformOperation.mockResolvedValueOnce({
        isValid: false,
        errors: [{ field: 'permission', message: 'Unauthorized' }]
      });

      await expect(service.createPolicy({
        PolicyName: 'Test'
      })).rejects.toThrow('Unauthorized');
    });

    it('should throw error if validation fails', async () => {
      await service.initialize();
      mockValidationServiceInstance.validatePolicyData.mockResolvedValueOnce({
        isValid: false,
        errors: [{ field: 'PolicyName', message: 'Name is required' }]
      });

      await expect(service.createPolicy({
        PolicyName: ''
      })).rejects.toThrow('Name is required');
    });

    it('should generate policy number if not provided', async () => {
      await service.initialize();

      const mockCreatedPolicy = createMockPolicy();
      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(mockCreatedPolicy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.createPolicy({
        PolicyName: 'Test Policy',
        PolicyCategory: PolicyCategory.ITSecurity
      });

      const addCall = mockSP._mockList.items.add.mock.calls[0][0];
      expect(addCall.PolicyNumber).toBeDefined();
    });
  });

  // ==========================================
  // getPolicyById Tests
  // ==========================================

  describe('getPolicyById', () => {
    it('should return cached policy if available', async () => {
      const cachedPolicy = createMockPolicy();
      mockCacheService.getPolicy.mockReturnValueOnce(cachedPolicy);

      const result = await service.getPolicyById(1);

      expect(result).toEqual(cachedPolicy);
      expect(mockSP._mockList.items.getById).not.toHaveBeenCalled();
    });

    it('should fetch from SharePoint if not cached', async () => {
      const spPolicy = createMockPolicy();
      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(spPolicy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      const result = await service.getPolicyById(1);

      expect(getByIdMock).toHaveBeenCalledWith(1);
      expect(mockCacheService.setPolicy).toHaveBeenCalled();
    });

    it('should bypass cache when requested', async () => {
      const cachedPolicy = createMockPolicy({ PolicyName: 'Cached' });
      mockCacheService.getPolicy.mockReturnValueOnce(cachedPolicy);

      const spPolicy = createMockPolicy({ PolicyName: 'Fresh' });
      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(spPolicy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      const result = await service.getPolicyById(1, true);

      expect(getByIdMock).toHaveBeenCalled();
    });
  });

  // ==========================================
  // updatePolicy Tests
  // ==========================================

  describe('updatePolicy', () => {
    it('should update policy and invalidate cache', async () => {
      await service.initialize();

      const updatedPolicy = createMockPolicy({ PolicyName: 'Updated' });
      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(updatedPolicy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      const result = await service.updatePolicy(1, { PolicyName: 'Updated' });

      expect(updateMock).toHaveBeenCalled();
      expect(mockCacheService.invalidatePolicy).toHaveBeenCalledWith(1);
    });

    it('should handle array fields as JSON', async () => {
      await service.initialize();

      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(createMockPolicy())
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.updatePolicy(1, {
        Tags: ['tag1', 'tag2'],
        RelatedPolicyIds: [2, 3]
      });

      const updateCall = updateMock.mock.calls[0][0];
      expect(updateCall.Tags).toBe(JSON.stringify(['tag1', 'tag2']));
      expect(updateCall.RelatedPolicyIds).toBe(JSON.stringify([2, 3]));
    });
  });

  // ==========================================
  // deletePolicy Tests
  // ==========================================

  describe('deletePolicy', () => {
    it('should archive policy instead of hard delete', async () => {
      await service.initialize();

      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(createMockPolicy())
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.deletePolicy(1);

      const updateCall = updateMock.mock.calls[0][0];
      expect(updateCall.Status).toBe(PolicyStatus.Archived);
      expect(updateCall.IsActive).toBe(false);
    });
  });

  // ==========================================
  // getPolicies Tests
  // ==========================================

  describe('getPolicies', () => {
    it('should return cached list if available', async () => {
      const cachedPolicies = [createMockPolicy()];
      mockCacheService.getPolicyList.mockReturnValueOnce(cachedPolicies);

      const result = await service.getPolicies();

      expect(result).toEqual(cachedPolicies);
    });

    it('should fetch from SharePoint with filters', async () => {
      const policies = [createMockPolicy()];
      const topMock = jest.fn().mockResolvedValue(policies);
      const filterMock = jest.fn().mockReturnValue({ top: topMock });

      mockSP._mockList.items.select = jest.fn().mockReturnThis();
      mockSP._mockList.items.expand = jest.fn().mockReturnValue({
        filter: filterMock,
        top: topMock
      });

      const result = await service.getPolicies({
        status: PolicyStatus.Published,
        category: PolicyCategory.HRPolicies
      });

      expect(filterMock).toHaveBeenCalled();
    });

    it('should cache results after fetching', async () => {
      const policies = [createMockPolicy()];
      const topMock = jest.fn().mockResolvedValue(policies);

      mockSP._mockList.items.select = jest.fn().mockReturnThis();
      mockSP._mockList.items.expand = jest.fn().mockReturnValue({ top: topMock });

      await service.getPolicies();

      expect(mockCacheService.setPolicyList).toHaveBeenCalled();
    });
  });

  // ==========================================
  // getPoliciesPaginated Tests
  // ==========================================

  describe('getPoliciesPaginated', () => {
    it('should return paginated results', async () => {
      const policies = Array.from({ length: 25 }, (_, i) =>
        createMockPolicy({ Id: i + 1, PolicyName: `Policy ${i + 1}` })
      );
      mockCacheService.getPolicyList.mockReturnValueOnce(policies);

      const result = await service.getPoliciesPaginated(1, 10);

      expect(result.items).toHaveLength(10);
      expect(result.totalCount).toBe(25);
      expect(result.totalPages).toBe(3);
      expect(result.hasNextPage).toBe(true);
      expect(result.hasPreviousPage).toBe(false);
    });

    it('should filter by search term', async () => {
      const policies = [
        createMockPolicy({ Id: 1, PolicyName: 'HR Policy' }),
        createMockPolicy({ Id: 2, PolicyName: 'IT Security Policy' }),
        createMockPolicy({ Id: 3, PolicyName: 'Finance Policy' })
      ];
      mockCacheService.getPolicyList.mockReturnValueOnce(policies);

      const result = await service.getPoliciesPaginated(1, 10, {
        searchTerm: 'Security'
      });

      expect(result.items).toHaveLength(1);
      expect(result.items[0].PolicyName).toBe('IT Security Policy');
    });

    it('should sort results', async () => {
      const policies = [
        createMockPolicy({ Id: 1, PolicyName: 'Zebra Policy' }),
        createMockPolicy({ Id: 2, PolicyName: 'Alpha Policy' }),
        createMockPolicy({ Id: 3, PolicyName: 'Beta Policy' })
      ];
      mockCacheService.getPolicyList.mockReturnValueOnce(policies);

      const result = await service.getPoliciesPaginated(1, 10, {
        sortBy: 'PolicyName',
        sortDirection: 'asc'
      });

      expect(result.items[0].PolicyName).toBe('Alpha Policy');
      expect(result.items[2].PolicyName).toBe('Zebra Policy');
    });
  });

  // ==========================================
  // Cache Management Tests
  // ==========================================

  describe('Cache Management', () => {
    it('should return cache stats', () => {
      mockCacheService.getStats.mockReturnValue({
        hits: 10,
        misses: 5,
        size: 15,
        hitRate: 66.67
      });

      const stats = service.getCacheStats();

      expect(stats.hits).toBe(10);
      expect(stats.misses).toBe(5);
    });

    it('should clear cache', () => {
      service.clearCache();

      expect(mockCacheService.clear).toHaveBeenCalled();
    });

    it('should refresh policy cache', async () => {
      const freshPolicy = createMockPolicy({ PolicyName: 'Fresh' });
      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(freshPolicy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.refreshPolicyCache(1);

      expect(mockCacheService.invalidatePolicy).toHaveBeenCalledWith(1);
    });
  });

  // ==========================================
  // Policy Lifecycle Tests
  // ==========================================

  describe('Policy Lifecycle', () => {
    beforeEach(async () => {
      await service.initialize();
    });

    it('should submit policy for review', async () => {
      const policy = createMockPolicy({ Status: PolicyStatus.Draft });
      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(policy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.submitForReview(1, [2, 3]);

      const updateCall = updateMock.mock.calls[0][0];
      expect(updateCall.Status).toBe(PolicyStatus.InReview);
      expect(updateCall.ReviewerIds).toEqual([2, 3]);
    });

    it('should approve policy when user has permission', async () => {
      const policy = createMockPolicy({ Status: PolicyStatus.InReview });
      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(policy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.approvePolicy(1, 'Approved');

      // Note: auditService.logPolicyApproval is called internally
      expect(updateMock).toHaveBeenCalled();
    });

    it('should reject approval for wrong status', async () => {
      const policy = createMockPolicy({ Status: PolicyStatus.Draft });
      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(policy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await expect(service.approvePolicy(1)).rejects.toThrow(
        'Cannot approve policy in Draft status'
      );
    });

    it('should reject policy with reason', async () => {
      const policy = createMockPolicy({ Status: PolicyStatus.InReview });
      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(policy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.rejectPolicy(1, 'Needs more detail');

      // Audit logging is tested via integration - mock is inline in jest.mock factory
      expect(updateMock).toHaveBeenCalled();
    });

    it('should require rejection reason', async () => {
      await expect(service.rejectPolicy(1, '')).rejects.toThrow(
        'Rejection reason is required'
      );
    });

    it('should archive policy', async () => {
      const policy = createMockPolicy();
      const updateMock = jest.fn().mockResolvedValue(undefined);
      const getByIdMock = jest.fn().mockReturnValue({
        update: updateMock,
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(policy)
      });
      mockSP._mockList.items.getById = getByIdMock;

      await service.archivePolicy(1, 'No longer needed');

      const updateCall = updateMock.mock.calls[0][0];
      expect(updateCall.Status).toBe(PolicyStatus.Archived);
    });
  });

  // ==========================================
  // Exemption Tests
  // ==========================================

  describe('Exemptions', () => {
    beforeEach(async () => {
      await service.initialize();
    });

    it('should request exemption', async () => {
      const result = await service.requestExemption({
        PolicyId: 1,
        UserId: 2,
        ExemptionReason: 'On leave',
        ExemptionType: 'Temporary'
      });

      expect(mockSP._mockList.items.add).toHaveBeenCalled();
    });

    it('should approve exemption', async () => {
      const exemption = { Id: 1, PolicyId: 1, UserId: 2 };
      const updateMock = jest.fn().mockResolvedValue(undefined);
      const filterMock = jest.fn().mockReturnValue({
        orderBy: jest.fn().mockReturnValue({
          top: jest.fn().mockResolvedValue([])
        })
      });

      mockSP._mockList.items.getById = jest.fn().mockReturnValue({
        update: updateMock,
        __call: jest.fn().mockResolvedValue(exemption)
      });
      mockSP._mockList.items.filter = filterMock;

      await service.approveExemption(1, 'Approved temporarily');

      expect(updateMock).toHaveBeenCalled();
    });
  });

  // ==========================================
  // Dashboard Tests
  // ==========================================

  describe('Dashboards', () => {
    it('should get user dashboard', async () => {
      const acknowledgements = [
        { Status: AcknowledgementStatus.Sent },
        { Status: AcknowledgementStatus.Acknowledged },
        { Status: AcknowledgementStatus.Overdue }
      ];

      const topMock = jest.fn().mockResolvedValue(acknowledgements);
      mockSP._mockList.items.filter = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnValue({
          top: topMock
        })
      });

      const result = await service.getUserDashboard(1);

      expect(result.totalPending).toBe(1);
      expect(result.totalOverdue).toBe(1);
      expect(result.totalCompleted).toBe(1);
    });

    it('should get policy compliance summary', async () => {
      const policy = createMockPolicy();
      const acknowledgements = [
        { Status: AcknowledgementStatus.Acknowledged, AssignedDate: '2024-01-01', AcknowledgedDate: '2024-01-05' },
        { Status: AcknowledgementStatus.Acknowledged, AssignedDate: '2024-01-01', AcknowledgedDate: '2024-01-03' },
        { Status: AcknowledgementStatus.Overdue }
      ];

      const getByIdMock = jest.fn().mockReturnValue({
        select: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        __call: jest.fn().mockResolvedValue(policy)
      });
      mockSP._mockList.items.getById = getByIdMock;
      mockSP._mockList.items.filter = jest.fn().mockReturnValue({
        top: jest.fn().mockResolvedValue(acknowledgements)
      });

      const result = await service.getPolicyComplianceSummary(1);

      expect(result.totalAssigned).toBe(3);
      expect(result.totalAcknowledged).toBe(2);
      expect(result.totalOverdue).toBe(1);
    });

    it('should get dashboard metrics', async () => {
      const policies = [
        createMockPolicy({ IsActive: true, Status: PolicyStatus.Published }),
        createMockPolicy({ IsActive: false, Status: PolicyStatus.Draft }),
        createMockPolicy({ IsActive: true, Status: PolicyStatus.Published, ComplianceRisk: ComplianceRisk.Critical })
      ];
      mockCacheService.getPolicyList.mockReturnValueOnce(policies);

      const acknowledgements = [
        { Status: AcknowledgementStatus.Acknowledged },
        { Status: AcknowledgementStatus.Overdue }
      ];
      mockSP._mockList.items.top = jest.fn().mockResolvedValue(acknowledgements);

      const result = await service.getDashboardMetrics();

      expect(result.totalPolicies).toBe(3);
      expect(result.activePolicies).toBe(2);
      expect(result.draftPolicies).toBe(1);
      expect(result.criticalRiskPolicies).toBe(1);
    });
  });
});
