// @ts-nocheck
// DataPrivacyService Unit Tests
// Testing GDPR compliance, data retention, anonymization, and deletion workflows

import { DataPrivacyService, IDataPrivacyConfig } from './DataPrivacyService';
import { SPFI } from '@pnp/sp';
import {
  EntityType,
  DeletionRequestType,
  DeletionRequestStatus,
  ExportFormat,
  ExportRequestStatus,
  AnonymizationMethod,
  ConsentType,
  PersonalDataType,
  AuditAction
} from '../models/IDataPrivacy';

// Mock PnP SP
jest.mock('@pnp/sp');

describe('DataPrivacyService', () => {
  let service: DataPrivacyService;
  let mockSp: jest.Mocked<SPFI>;
  let mockConfig: IDataPrivacyConfig;

  beforeEach(() => {
    // Setup mock SharePoint instance
    mockSp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnThis(),
          items: {
            filter: jest.fn().mockReturnThis(),
            top: jest.fn().mockReturnThis(),
            orderBy: jest.fn().mockResolvedValue([]),
            getById: jest.fn().mockReturnThis(),
            add: jest.fn().mockResolvedValue({ data: { Id: 1 } }),
            update: jest.fn().mockResolvedValue({}),
            delete: jest.fn().mockResolvedValue({})
          }
        }
      }
    } as any;

    mockConfig = {
      sp: mockSp,
      currentUserEmail: 'test@company.com',
      currentUserId: 'test-user-id',
      encryptionKey: 'test-encryption-key-minimum-32-characters-long-secure'
    };

    service = new DataPrivacyService(mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  // ==========================================
  // CRITICAL: Security Tests
  // ==========================================

  describe('Security - Encryption Key Validation', () => {
    test('CRITICAL: should throw error when encryption key is missing', () => {
      const invalidConfig = { ...mockConfig, encryptionKey: undefined };

      expect(() => new DataPrivacyService(invalidConfig as any))
        .toThrow('Encryption key is required');
    });

    test('CRITICAL: should throw error when encryption key is too short', () => {
      const invalidConfig = { ...mockConfig, encryptionKey: 'short' };

      expect(() => new DataPrivacyService(invalidConfig))
        .toThrow('Encryption key must be at least 32 characters');
    });

    test('CRITICAL: should accept valid encryption key', () => {
      expect(() => new DataPrivacyService(mockConfig)).not.toThrow();
    });
  });

  describe('Security - SQL Injection Prevention', () => {
    test('CRITICAL: should sanitize user IDs in filter queries', async () => {
      const maliciousUserId = "'; DROP TABLE PM_Processes; --";

      await service.initialize();

      // Should not execute the malicious SQL
      await expect(async () => {
        await service.getUserDataItems('PM_Processes', maliciousUserId);
      }).not.toThrow();

      // Verify filter was called with sanitized input
      const filterCall = mockSp.web.lists.getByTitle().items.filter;
      expect(filterCall).toHaveBeenCalled();

      // Filter should NOT contain the raw malicious input
      const filterArg = (filterCall as jest.Mock).mock.calls[0][0];
      expect(filterArg).not.toContain("DROP TABLE");
    });

    test('CRITICAL: should validate email addresses before using in filters', async () => {
      const maliciousEmail = "test@company.com' OR '1'='1";

      await service.initialize();

      // Should throw validation error before executing query
      await expect(
        service.getUserConsents(maliciousEmail)
      ).rejects.toThrow('Invalid email format');
    });

    test('CRITICAL: should validate entity types to prevent injection', async () => {
      const maliciousEntityType = "Process'; DELETE FROM PM_AuditLog; --" as EntityType;

      await service.initialize();

      // Should validate entity type against enum
      await expect(
        service.getRetentionPolicyByEntity(maliciousEntityType)
      ).rejects.toThrow('Invalid entity type');
    });
  });

  describe('Security - Input Validation', () => {
    test('should validate email format', () => {
      const invalidEmails = [
        '',
        'not-an-email',
        'missing@domain',
        '@nodomain.com',
        'spaces in@email.com',
        'double@@email.com'
      ];

      invalidEmails.forEach(email => {
        expect(() => service.validateEmail(email)).toThrow('Invalid email format');
      });
    });

    test('should accept valid email formats', () => {
      const validEmails = [
        'user@company.com',
        'first.last@company.co.uk',
        'user+tag@domain.org'
      ];

      validEmails.forEach(email => {
        expect(() => service.validateEmail(email)).not.toThrow();
      });
    });

    test('should validate date ranges', () => {
      const futureDate = new Date('2099-01-01');
      const pastDate = new Date('1900-01-01');

      expect(() => service.validateDateRange(futureDate, pastDate))
        .toThrow('End date must be after start date');
    });
  });

  // ==========================================
  // Initialization Tests
  // ==========================================

  describe('Initialization', () => {
    test('should verify all required lists exist', async () => {
      const requiredLists = [
        'PM_DataRetentionPolicies',
        'PM_DataDeletionRequests',
        'PM_DataExportRequests',
        'PM_ConsentRecords',
        'PM_PrivacyImpactAssessments',
        'PM_AnonymizationJobs',
        'PM_AuditLog',
        'PM_DataSubjectRequests'
      ];

      await service.initialize();

      requiredLists.forEach(listName => {
        expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith(listName);
      });
    });

    test('should throw error if required list is missing', async () => {
      mockSp.web.lists.getByTitle = jest.fn().mockRejectedValue(
        new Error('List not found')
      );

      await expect(service.initialize()).rejects.toThrow(
        'DataPrivacyService initialization failed'
      );
    });
  });

  // ==========================================
  // Data Retention Policy Tests
  // ==========================================

  describe('Data Retention Policies', () => {
    test('should retrieve all active retention policies', async () => {
      const mockPolicies = [
        { Id: 1, EntityType: 'Process', RetentionPeriodDays: 2555, IsActive: true },
        { Id: 2, EntityType: 'Task', RetentionPeriodDays: 1095, IsActive: true }
      ];

      mockSp.web.lists.getByTitle().items.orderBy = jest.fn().mockResolvedValue(mockPolicies);

      const policies = await service.getRetentionPolicies();

      expect(policies).toHaveLength(2);
      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('PM_DataRetentionPolicies');
      expect(mockSp.web.lists.getByTitle().items.filter).toHaveBeenCalledWith('IsActive eq true');
    });

    test('should get retention policy for specific entity type', async () => {
      const mockPolicy = {
        Id: 1,
        EntityType: 'Process',
        RetentionPeriodDays: 2555,
        IsActive: true
      };

      mockSp.web.lists.getByTitle().items.top = jest.fn().mockResolvedValue([mockPolicy]);

      const policy = await service.getRetentionPolicyByEntity(EntityType.Process);

      expect(policy).toBeDefined();
      expect(policy?.EntityType).toBe('Process');
      expect(mockSp.web.lists.getByTitle().items.filter).toHaveBeenCalledWith(
        expect.stringContaining("EntityType eq 'Process'")
      );
    });

    test('should return null if no policy exists for entity type', async () => {
      mockSp.web.lists.getByTitle().items.top = jest.fn().mockResolvedValue([]);

      const policy = await service.getRetentionPolicyByEntity(EntityType.Task);

      expect(policy).toBeNull();
    });

    test('should execute retention policies with auto-delete enabled', async () => {
      const mockPolicies = [
        {
          Id: 1,
          EntityType: 'Process',
          RetentionPeriodDays: 7,
          AutoDeleteEnabled: true,
          AnonymizeBeforeDelete: true
        }
      ];

      mockSp.web.lists.getByTitle().items.filter = jest.fn()
        .mockResolvedValueOnce(mockPolicies) // For getRetentionPolicies
        .mockResolvedValueOnce([{ Id: 100 }, { Id: 101 }]); // For old items

      const results = await service.executeRetentionPolicies();

      expect(results['Process']).toBeDefined();
      expect(results['Process'].recordsProcessed).toBe(2);
    });
  });

  // ==========================================
  // Anonymization Tests
  // ==========================================

  describe('Data Anonymization', () => {
    test('should anonymize item with HASH method', async () => {
      const result = await service.anonymizeItem('PM_Processes', 1, EntityType.Process);

      expect(result).toBe(true);
      expect(mockSp.web.lists.getByTitle().items.getById).toHaveBeenCalledWith(1);
      expect(mockSp.web.lists.getByTitle().items.update).toHaveBeenCalled();
    });

    test('should hash values using SHA-256', () => {
      const value = 'test@company.com';
      const hashed = service.anonymizeValue(value, AnonymizationMethod.Hash, PersonalDataType.Email);

      // SHA-256 hash should be 64 characters
      expect(hashed).toHaveLength(64);
      // Should be deterministic
      const hashed2 = service.anonymizeValue(value, AnonymizationMethod.Hash, PersonalDataType.Email);
      expect(hashed).toBe(hashed2);
    });

    test('should mask email addresses correctly', () => {
      const email = 'john.smith@company.com';
      const masked = service.anonymizeValue(email, AnonymizationMethod.Mask, PersonalDataType.Email);

      expect(masked).toMatch(/^jo\*\*\*@company\.com$/);
    });

    test('should mask phone numbers correctly', () => {
      const phone = '555-123-4567';
      const masked = service.anonymizeValue(phone, AnonymizationMethod.Mask, PersonalDataType.Phone);

      expect(masked).toMatch(/\*\*\*-\*\*\*-4567$/);
    });

    test('should replace names with generic placeholder', () => {
      const name = 'John Smith';
      const replaced = service.anonymizeValue(name, AnonymizationMethod.Replace, PersonalDataType.Name);

      expect(replaced).toBe('[Anonymized User]');
    });

    test('should encrypt values using AES', () => {
      const value = 'sensitive-data';
      const encrypted = service.anonymizeValue(value, AnonymizationMethod.Encrypt, PersonalDataType.Other);

      expect(encrypted).not.toBe(value);
      expect(encrypted.length).toBeGreaterThan(0);
      // Should be different each time (includes IV)
      const encrypted2 = service.anonymizeValue(value, AnonymizationMethod.Encrypt, PersonalDataType.Other);
      expect(encrypted).not.toBe(encrypted2);
    });

    test('should create anonymization job', async () => {
      const jobId = await service.createAnonymizationJob(
        EntityType.Process,
        'user@company.com',
        new Date('2020-01-01'),
        new Date('2024-01-01')
      );

      expect(jobId).toBe(1);
      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('PM_AnonymizationJobs');
      expect(mockSp.web.lists.getByTitle().items.add).toHaveBeenCalled();
    });
  });

  // ==========================================
  // Right to be Forgotten Tests
  // ==========================================

  describe('Right to be Forgotten (RTBF)', () => {
    test('should submit deletion request', async () => {
      const requestId = await service.submitDeletionRequest(
        DeletionRequestType.FullDeletion,
        'user@company.com',
        'Employee leaving',
        [EntityType.Process, EntityType.Task]
      );

      expect(requestId).toBe(1);
      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('PM_DataDeletionRequests');

      const addCall = (mockSp.web.lists.getByTitle().items.add as jest.Mock).mock.calls[0][0];
      expect(addCall.RequestType).toBe(DeletionRequestType.FullDeletion);
      expect(addCall.Status).toBe(DeletionRequestStatus.Pending);
    });

    test('should process approved deletion request', async () => {
      const mockRequest = {
        Id: 1,
        RequestType: DeletionRequestType.FullDeletion,
        SubjectUserEmail: 'user@company.com',
        EntityTypes: [EntityType.Process],
        Status: DeletionRequestStatus.Pending
      };

      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockResolvedValue(mockRequest);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { Id: 1 }, { Id: 2 }, { Id: 3 }
      ]);

      const result = await service.processDeletionRequest(1, true);

      expect(result.success).toBe(true);
      expect(result.itemsDeleted).toBe(3);
      expect(mockSp.web.lists.getByTitle().items.update).toHaveBeenCalledWith(
        expect.objectContaining({ Status: DeletionRequestStatus.Completed })
      );
    });

    test('should process rejected deletion request', async () => {
      const mockRequest = {
        Id: 1,
        RequestType: DeletionRequestType.FullDeletion,
        Status: DeletionRequestStatus.Pending
      };

      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockResolvedValue(mockRequest);

      const result = await service.processDeletionRequest(1, false, 'Active employment');

      expect(result.success).toBe(true);
      expect(result.summary).toBe('Request rejected');
      expect(mockSp.web.lists.getByTitle().items.update).toHaveBeenCalledWith(
        expect.objectContaining({
          Status: DeletionRequestStatus.Rejected,
          RejectionReason: 'Active employment'
        })
      );
    });

    test('should anonymize instead of delete when request type is Anonymization', async () => {
      const mockRequest = {
        Id: 1,
        RequestType: DeletionRequestType.Anonymization,
        SubjectUserEmail: 'user@company.com',
        EntityTypes: [EntityType.Process]
      };

      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockResolvedValue(mockRequest);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { Id: 1 }, { Id: 2 }
      ]);

      const result = await service.processDeletionRequest(1, true);

      expect(result.itemsAnonymized).toBe(2);
      expect(result.itemsDeleted).toBe(0);
    });
  });

  // ==========================================
  // Data Export Tests
  // ==========================================

  describe('Data Export (Right to Data Portability)', () => {
    test('should submit export request', async () => {
      const requestId = await service.submitExportRequest(
        ExportFormat.JSON,
        true,
        [EntityType.Process, EntityType.Task]
      );

      expect(requestId).toBe(1);
      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('PM_DataExportRequests');

      const addCall = (mockSp.web.lists.getByTitle().items.add as jest.Mock).mock.calls[0][0];
      expect(addCall.ExportFormat).toBe(ExportFormat.JSON);
      expect(addCall.Status).toBe(ExportRequestStatus.Pending);
    });

    test('should generate JSON export content', () => {
      const data = {
        Process: [{ Id: 1, Title: 'Test Process' }],
        Task: [{ Id: 2, Title: 'Test Task' }]
      };

      const json = service.generateExportContent(data, ExportFormat.JSON);
      const parsed = JSON.parse(json);

      expect(parsed.Process).toHaveLength(1);
      expect(parsed.Task).toHaveLength(1);
    });

    test('should generate CSV export content', () => {
      const data = {
        Process: [
          { Id: 1, Title: 'Process 1', Status: 'Active' },
          { Id: 2, Title: 'Process 2', Status: 'Completed' }
        ]
      };

      const csv = service.generateExportContent(data, ExportFormat.CSV);

      expect(csv).toContain('=== Process ===');
      expect(csv).toContain('Id,Title,Status');
      expect(csv).toContain('1,Process 1,Active');
    });

    test('should generate XML export content', () => {
      const data = {
        Process: [{ Id: 1, Title: 'Test' }]
      };

      const xml = service.generateExportContent(data, ExportFormat.XML);

      expect(xml).toContain('<?xml version="1.0" encoding="UTF-8"?>');
      expect(xml).toContain('<DataExport>');
      expect(xml).toContain('<Process>');
      expect(xml).toContain('<Id>1</Id>');
    });

    test('should set 30-day expiry on export downloads', async () => {
      const mockRequest = {
        Id: 1,
        ExportFormat: ExportFormat.JSON,
        RequesterEmail: 'user@company.com',
        EntityTypes: [EntityType.Process]
      };

      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockResolvedValue(mockRequest);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([]);

      await service.processExportRequest(1);

      const updateCall = (mockSp.web.lists.getByTitle().items.update as jest.Mock).mock.calls[0][0];
      const expiryDate = new Date(updateCall.ExpiryDate);
      const expectedExpiry = new Date();
      expectedExpiry.setDate(expectedExpiry.getDate() + 30);

      // Should be approximately 30 days from now (allow 1 minute tolerance)
      const diff = Math.abs(expiryDate.getTime() - expectedExpiry.getTime());
      expect(diff).toBeLessThan(60000); // 1 minute
    });
  });

  // ==========================================
  // Consent Management Tests
  // ==========================================

  describe('Consent Management', () => {
    test('should record user consent', async () => {
      const consentId = await service.recordConsent(
        ConsentType.DataProcessing,
        'Employee lifecycle management',
        true,
        '1.0',
        'Web Form',
        '192.168.1.1',
        'Mozilla/5.0'
      );

      expect(consentId).toBe(1);

      const addCall = (mockSp.web.lists.getByTitle().items.add as jest.Mock).mock.calls[0][0];
      expect(addCall.ConsentGiven).toBe(true);
      expect(addCall.ConsentVersion).toBe('1.0');
      expect(addCall.IsActive).toBe(true);
    });

    test('should withdraw consent', async () => {
      await service.withdrawConsent(1, 'User request');

      expect(mockSp.web.lists.getByTitle().items.update).toHaveBeenCalledWith(
        expect.objectContaining({
          IsActive: false,
          WithdrawalReason: 'User request'
        })
      );
    });

    test('should retrieve user consents', async () => {
      const mockConsents = [
        { Id: 1, ConsentType: ConsentType.DataProcessing, IsActive: true },
        { Id: 2, ConsentType: ConsentType.Marketing, IsActive: false }
      ];

      mockSp.web.lists.getByTitle().items.orderBy = jest.fn().mockResolvedValue(mockConsents);

      const consents = await service.getUserConsents('user@company.com');

      expect(consents).toHaveLength(2);
      expect(mockSp.web.lists.getByTitle().items.filter).toHaveBeenCalledWith(
        expect.stringContaining("UserEmail eq 'user@company.com'")
      );
    });
  });

  // ==========================================
  // Privacy Impact Assessment Tests
  // ==========================================

  describe('Privacy Impact Assessments', () => {
    test('should create PIA with risk calculation', async () => {
      const risks = [
        { id: '1', riskScore: 9 }, // Critical
        { id: '2', riskScore: 5 }  // High
      ];

      const piaId = await service.createPIA({
        ProjectName: 'Test Project',
        Risks: risks
      });

      expect(piaId).toBe(1);

      const addCall = (mockSp.web.lists.getByTitle().items.add as jest.Mock).mock.calls[0][0];
      expect(addCall.RiskLevel).toBe('Critical'); // Based on max risk score of 9
    });

    test('should calculate risk level correctly', () => {
      expect(service.calculateOverallRiskLevel([{ riskScore: 9 }])).toBe('Critical');
      expect(service.calculateOverallRiskLevel([{ riskScore: 6 }])).toBe('High');
      expect(service.calculateOverallRiskLevel([{ riskScore: 4 }])).toBe('Medium');
      expect(service.calculateOverallRiskLevel([{ riskScore: 2 }])).toBe('Low');
      expect(service.calculateOverallRiskLevel([])).toBe('Low');
    });
  });

  // ==========================================
  // Audit Logging Tests
  // ==========================================

  describe('Audit Logging', () => {
    test('should log all GDPR actions', async () => {
      await service.logAuditEntry({
        Action: AuditAction.DataDeleted,
        EntityType: 'Process',
        EntityId: 123,
        Success: true
      });

      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('PM_AuditLog');

      const addCall = (mockSp.web.lists.getByTitle().items.add as jest.Mock).mock.calls[0][0];
      expect(addCall.Action).toBe(AuditAction.DataDeleted);
      expect(addCall.Success).toBe(true);
      expect(addCall.Timestamp).toBeDefined();
    });

    test('should include user information in audit logs', async () => {
      await service.logAuditEntry({
        Action: AuditAction.ConsentGiven,
        Success: true
      });

      const addCall = (mockSp.web.lists.getByTitle().items.add as jest.Mock).mock.calls[0][0];
      expect(addCall.UserId).toBe('test-user-id');
      expect(addCall.UserEmail).toBe('test@company.com');
    });
  });

  // ==========================================
  // Error Handling Tests
  // ==========================================

  describe('Error Handling', () => {
    test('should handle SharePoint errors gracefully', async () => {
      mockSp.web.lists.getByTitle = jest.fn().mockRejectedValue(
        new Error('SharePoint error')
      );

      await expect(service.getRetentionPolicies()).rejects.toThrow('SharePoint error');
    });

    test('should handle network timeouts', async () => {
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockRejectedValue(
        new Error('Network timeout')
      );

      await expect(service.getUserConsents()).rejects.toThrow('Network timeout');
    });

    test('should handle missing items gracefully', async () => {
      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockRejectedValue(
        new Error('Item not found')
      );

      const request = await service.getDeletionRequest(999);
      expect(request).toBeNull();
    });
  });

  // ==========================================
  // Performance Tests
  // ==========================================

  describe('Performance', () => {
    test('should use batch operations for multiple deletes', async () => {
      // This test verifies the service doesn't make N individual delete calls
      const mockRequest = {
        Id: 1,
        RequestType: DeletionRequestType.FullDeletion,
        SubjectUserEmail: 'user@company.com',
        EntityTypes: [EntityType.Process]
      };

      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockResolvedValue(mockRequest);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue(
        Array(100).fill({ Id: 1 })
      );

      await service.processDeletionRequest(1, true);

      // Should use batching, not 100 individual calls
      const deleteCallCount = (mockSp.web.lists.getByTitle().items.delete as jest.Mock).mock.calls.length;
      expect(deleteCallCount).toBeLessThan(100);
    });
  });
});
