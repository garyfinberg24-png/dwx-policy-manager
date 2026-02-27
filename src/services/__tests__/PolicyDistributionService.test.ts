/**
 * PolicyDistributionService Unit Tests
 *
 * Tests constructor, method signatures, and SP call behavior by mocking
 * the entire PnP SP fluent API chain.
 */

// ---------------------------------------------------------------------------
// Mock the LoggingService (imported by PolicyDistributionService)
// ---------------------------------------------------------------------------
jest.mock('../LoggingService', () => ({
  logger: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    debug: jest.fn(),
  },
}));

// ---------------------------------------------------------------------------
// Mock ValidationUtils
// ---------------------------------------------------------------------------
jest.mock('../../utils/ValidationUtils', () => ({
  ValidationUtils: {
    sanitizeForOData: jest.fn((val: string) => val),
  },
}));

// ---------------------------------------------------------------------------
// Mock SharePointListNames constants
// ---------------------------------------------------------------------------
jest.mock('../../constants/SharePointListNames', () => ({
  PolicyLists: {
    POLICY_DISTRIBUTIONS: 'PM_PolicyDistributions',
    POLICY_ACKNOWLEDGEMENTS: 'PM_PolicyAcknowledgements',
    POLICIES: 'PM_Policies',
  },
  PolicyPackLists: {
    POLICY_PACKS: 'PM_PolicyPacks',
  },
  SystemLists: {
    NOTIFICATION_QUEUE: 'PM_NotificationQueue',
  },
}));

import { PolicyDistributionService, ISPDistributionItem } from '../PolicyDistributionService';
import { logger } from '../LoggingService';
import { ValidationUtils } from '../../utils/ValidationUtils';

// ---------------------------------------------------------------------------
// PnP SP mock builder
// ---------------------------------------------------------------------------

/** Build a mock that mimics the PnP fluent API: sp.web.lists.getByTitle(...).items... */
function createMockSp(resolvedValue: unknown = [], addResult?: unknown) {
  const itemsMock = {
    select: jest.fn().mockReturnThis(),
    expand: jest.fn().mockReturnThis(),
    filter: jest.fn().mockReturnThis(),
    orderBy: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolvedValue)),
    // PnP v3 .items.add() returns IItemAddResult { data: {...} }
    add: jest.fn().mockResolvedValue(addResult ?? { data: resolvedValue }),
    getById: jest.fn().mockReturnValue({
      update: jest.fn().mockResolvedValue(undefined),
      delete: jest.fn().mockResolvedValue(undefined),
    }),
  };

  const sp = {
    web: {
      lists: {
        getByTitle: jest.fn().mockReturnValue({
          items: itemsMock,
        }),
      },
    },
  };

  return { sp, itemsMock };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('PolicyDistributionService', () => {
  let service: PolicyDistributionService;
  let mockSp: ReturnType<typeof createMockSp>;

  beforeEach(() => {
    jest.clearAllMocks();
    mockSp = createMockSp();
    service = new PolicyDistributionService(mockSp.sp as any);
  });

  // ===== Constructor =====

  describe('constructor', () => {
    it('should create an instance without throwing', () => {
      expect(service).toBeDefined();
    });
  });

  // ===== getDistributions =====

  describe('getDistributions', () => {
    it('should call the correct SharePoint list', async () => {
      await service.getDistributions();
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_PolicyDistributions');
    });

    it('should return distribution items from SharePoint', async () => {
      const mockData: Partial<ISPDistributionItem>[] = [
        { Id: 1, Title: 'Campaign 1', DistributionName: 'Jan Campaign', TargetCount: 100, TotalSent: 50 },
        { Id: 2, Title: 'Campaign 2', DistributionName: 'Feb Campaign', TargetCount: 200, TotalSent: 180 },
      ];
      const { sp } = createMockSp(mockData);
      service = new PolicyDistributionService(sp as any);

      const result = await service.getDistributions();
      expect(result).toHaveLength(2);
      expect(result[0].Id).toBe(1);
    });

    it('should select the expected fields', async () => {
      await service.getDistributions();
      expect(mockSp.itemsMock.select).toHaveBeenCalled();
      const selectArgs = mockSp.itemsMock.select.mock.calls[0];
      expect(selectArgs).toContain('Id');
      expect(selectArgs).toContain('DistributionName');
      expect(selectArgs).toContain('Status');
    });

    it('should order by Modified descending', async () => {
      await service.getDistributions();
      expect(mockSp.itemsMock.orderBy).toHaveBeenCalledWith('Modified', false);
    });

    it('should log info on success', async () => {
      await service.getDistributions();
      expect(logger.info).toHaveBeenCalledWith(
        'PolicyDistributionService',
        expect.stringContaining('distributions from SharePoint')
      );
    });

    it('should throw and log error on failure', async () => {
      const error = new Error('SP Error');
      const failSp = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                select: jest.fn().mockReturnThis(),
                expand: jest.fn().mockReturnThis(),
                orderBy: jest.fn().mockReturnThis(),
                top: jest.fn().mockReturnValue(jest.fn().mockRejectedValue(error)),
              },
            }),
          },
        },
      };
      service = new PolicyDistributionService(failSp as any);
      await expect(service.getDistributions()).rejects.toThrow('SP Error');
      expect(logger.error).toHaveBeenCalledWith(
        'PolicyDistributionService',
        'getDistributions failed:',
        error
      );
    });
  });

  // ===== createDistribution =====

  describe('createDistribution', () => {
    it('should add a new item to the distributions list', async () => {
      const data = { Title: 'New Campaign', DistributionScope: 'All' };
      const addData = { Id: 99, ...data };
      const { sp, itemsMock } = createMockSp([], { data: addData });
      service = new PolicyDistributionService(sp as any);

      const result = await service.createDistribution(data);
      expect(itemsMock.add).toHaveBeenCalledWith(data);
      expect(result.Id).toBe(99);
    });

    it('should log info on successful creation', async () => {
      const { sp } = createMockSp([], { data: { Id: 42 } });
      service = new PolicyDistributionService(sp as any);
      await service.createDistribution({ Title: 'Test' });
      expect(logger.info).toHaveBeenCalledWith(
        'PolicyDistributionService',
        expect.stringContaining('Created distribution')
      );
    });
  });

  // ===== updateDistribution =====

  describe('updateDistribution', () => {
    it('should update the item by ID', async () => {
      const data = { Status: 'Active' };
      await service.updateDistribution(5, data);
      expect(mockSp.itemsMock.getById).toHaveBeenCalledWith(5);
    });

    it('should log info on successful update', async () => {
      await service.updateDistribution(5, { Status: 'Active' });
      expect(logger.info).toHaveBeenCalledWith(
        'PolicyDistributionService',
        'Updated distribution id=5'
      );
    });
  });

  // ===== deleteDistribution =====

  describe('deleteDistribution', () => {
    it('should delete the item by ID', async () => {
      await service.deleteDistribution(7);
      expect(mockSp.itemsMock.getById).toHaveBeenCalledWith(7);
    });

    it('should log info on successful deletion', async () => {
      await service.deleteDistribution(7);
      expect(logger.info).toHaveBeenCalledWith(
        'PolicyDistributionService',
        'Deleted distribution id=7'
      );
    });
  });

  // ===== getDistributionRecipients =====

  describe('getDistributionRecipients', () => {
    it('should query the acknowledgements list with sanitized ID', async () => {
      await service.getDistributionRecipients(42);
      expect(ValidationUtils.sanitizeForOData).toHaveBeenCalledWith('42');
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_PolicyAcknowledgements');
    });

    it('should filter by DistributionId', async () => {
      await service.getDistributionRecipients(42);
      expect(mockSp.itemsMock.filter).toHaveBeenCalledWith('DistributionId eq 42');
    });
  });

  // ===== getPolicies =====

  describe('getPolicies', () => {
    it('should query PM_Policies with Published filter', async () => {
      await service.getPolicies();
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_Policies');
      expect(mockSp.itemsMock.filter).toHaveBeenCalledWith("PolicyStatus eq 'Published'");
    });
  });

  // ===== getPolicyPacks =====

  describe('getPolicyPacks', () => {
    it('should query PM_PolicyPacks', async () => {
      await service.getPolicyPacks();
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_PolicyPacks');
    });
  });

  // ===== calculateCampaignMetrics =====

  describe('calculateCampaignMetrics', () => {
    it('should query acknowledgements list with sanitized distribution ID', async () => {
      await service.calculateCampaignMetrics(10);
      expect(ValidationUtils.sanitizeForOData).toHaveBeenCalledWith('10');
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_PolicyAcknowledgements');
    });

    it('should calculate metrics from recipient data', async () => {
      const mockRecipients = [
        { Id: 1, AckStatus: 'Acknowledged', SentDate: '2026-01-01', OpenedDate: '2026-01-02', AcknowledgedDate: '2026-01-03', DueDate: '2027-06-01' },
        { Id: 2, AckStatus: 'Opened', SentDate: '2026-01-01', OpenedDate: '2026-01-05', DueDate: '2027-06-01' },
        { Id: 3, AckStatus: 'Pending', SentDate: '2026-01-01', DueDate: '2025-01-01' }, // overdue (past)
        { Id: 4, AckStatus: 'Failed', DueDate: '2027-06-01' },
      ];
      const { sp } = createMockSp(mockRecipients);
      service = new PolicyDistributionService(sp as any);

      const metrics = await service.calculateCampaignMetrics(10);
      expect(metrics.totalSent).toBe(3);         // 3 have SentDate
      expect(metrics.totalDelivered).toBe(3);     // 3 sent - 0 failed among sent
      expect(metrics.totalAcknowledged).toBe(1);  // 1 Acknowledged
      expect(metrics.totalOverdue).toBe(1);       // 1 past due and not ack'd
      expect(metrics.totalFailed).toBe(1);        // 1 Failed
      expect(metrics.ackRate).toBe(33);           // 1/3 = 33%
    });

    it('should return zero metrics when no recipients exist', async () => {
      const { sp } = createMockSp([]);
      service = new PolicyDistributionService(sp as any);

      const metrics = await service.calculateCampaignMetrics(99);
      expect(metrics.totalSent).toBe(0);
      expect(metrics.ackRate).toBe(0);
    });

    it('should log metrics summary', async () => {
      await service.calculateCampaignMetrics(5);
      expect(logger.info).toHaveBeenCalledWith(
        'PolicyDistributionService',
        expect.stringContaining('Calculated metrics for distribution id=5')
      );
    });
  });

  // ===== sendEscalationNotifications =====

  describe('sendEscalationNotifications', () => {
    it('should create notification queue entries for overdue recipients', async () => {
      const overdueRecipients = [
        { Id: 1, Title: 'John Doe', UserEmail: 'john@example.com', AckStatus: 'Pending', DueDate: '2025-12-01' },
        { Id: 2, Title: 'Jane Smith', UserEmail: 'jane@example.com', AckStatus: 'Pending', DueDate: '2025-12-01' },
      ];

      const queued = await service.sendEscalationNotifications(10, 'Test Campaign', overdueRecipients);
      expect(queued).toBe(2);
      expect(mockSp.sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_NotificationQueue');
      expect(mockSp.itemsMock.add).toHaveBeenCalledTimes(2);
    });

    it('should return 0 when no recipients provided', async () => {
      const queued = await service.sendEscalationNotifications(10, 'Test Campaign', []);
      expect(queued).toBe(0);
    });

    it('should log the count of queued notifications', async () => {
      await service.sendEscalationNotifications(10, 'Test Campaign', [
        { Id: 1, Title: 'User', UserEmail: 'user@example.com', AckStatus: 'Pending' },
      ]);
      expect(logger.info).toHaveBeenCalledWith(
        'PolicyDistributionService',
        expect.stringContaining('Queued 1/1 escalation notifications')
      );
    });
  });
});
