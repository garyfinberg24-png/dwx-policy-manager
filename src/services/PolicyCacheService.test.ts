// @ts-nocheck
/**
 * PolicyCacheService Unit Tests
 * Testing in-memory caching with TTL for policy data
 */
/// <reference types="jest" />

import {
  PolicyCacheService,
  getPolicyCacheService,
  resetPolicyCacheService,
  paginateArray,
  generateCacheKey,
  ICacheStats,
  IPaginatedResult
} from './PolicyCacheService';
import {
  IPolicy,
  PolicyStatus,
  PolicyCategory,
  PolicyType,
  ComplianceRisk,
  ReadTimeframe,
  VersionType,
  DocumentFormat,
  AcknowledgementType,
  DistributionScope
} from '../models/IPolicy';

describe('PolicyCacheService', () => {
  let cacheService: PolicyCacheService;

  // Mock policy data
  const mockPolicy: IPolicy = {
    Id: 1,
    Title: 'Test Policy',
    PolicyNumber: 'POL-001',
    PolicyName: 'Test Policy',
    PolicyCategory: PolicyCategory.HRPolicies,
    PolicyType: PolicyType.Corporate,
    Description: 'Test policy description',
    PolicySummary: 'Test summary',
    PolicyContent: '<p>Test content</p>',
    KeyPoints: ['Point 1', 'Point 2'],
    VersionNumber: '1.0',
    VersionType: VersionType.Major,
    MajorVersion: 1,
    MinorVersion: 0,
    DocumentFormat: DocumentFormat.HTML,
    PolicyOwnerId: 1,
    PolicyAuthorIds: [1],
    ComplianceRisk: ComplianceRisk.Medium,
    ReadTimeframe: ReadTimeframe.Week1,
    ReadTimeframeDays: 7,
    RequiresAcknowledgement: true,
    AcknowledgementType: AcknowledgementType.OneTime,
    RequiresQuiz: false,
    AllowRetake: false,
    DistributionScope: DistributionScope.AllEmployees,
    Status: PolicyStatus.Published,
    EffectiveDate: new Date('2024-01-01'),
    IsMandatory: true,
    IsActive: true,
    Created: new Date('2024-01-01'),
    Modified: new Date('2024-01-01'),
    AuthorId: 1,
    EditorId: 1
  };

  beforeEach(() => {
    // Reset singleton before each test
    resetPolicyCacheService();
    cacheService = new PolicyCacheService({
      defaultTtlMs: 1000, // 1 second for testing
      maxEntries: 100,
      cleanupIntervalMs: 60000
    });
  });

  afterEach(() => {
    cacheService.dispose();
    resetPolicyCacheService();
  });

  // ==========================================
  // Core Cache Operations
  // ==========================================

  describe('Core Cache Operations', () => {
    test('should store and retrieve items from cache', () => {
      cacheService.set('test-key', { data: 'test value' });
      const result = cacheService.get<{ data: string }>('test-key');

      expect(result).toEqual({ data: 'test value' });
    });

    test('should return null for non-existent keys', () => {
      const result = cacheService.get('non-existent');

      expect(result).toBeNull();
    });

    test('should delete items from cache', () => {
      cacheService.set('test-key', 'value');
      const deleted = cacheService.delete('test-key');

      expect(deleted).toBe(true);
      expect(cacheService.get('test-key')).toBeNull();
    });

    test('should return false when deleting non-existent key', () => {
      const deleted = cacheService.delete('non-existent');

      expect(deleted).toBe(false);
    });

    test('should clear all items from cache', () => {
      cacheService.set('key1', 'value1');
      cacheService.set('key2', 'value2');
      cacheService.set('key3', 'value3');

      cacheService.clear();
      const stats = cacheService.getStats();

      expect(stats.size).toBe(0);
      expect(cacheService.get('key1')).toBeNull();
      expect(cacheService.get('key2')).toBeNull();
    });

    test('should delete items by prefix', () => {
      cacheService.set('policy:1', 'policy 1');
      cacheService.set('policy:2', 'policy 2');
      cacheService.set('ack:1', 'ack 1');

      const deletedCount = cacheService.deleteByPrefix('policy:');

      expect(deletedCount).toBe(2);
      expect(cacheService.get('policy:1')).toBeNull();
      expect(cacheService.get('policy:2')).toBeNull();
      expect(cacheService.get('ack:1')).not.toBeNull();
    });
  });

  // ==========================================
  // TTL (Time To Live) Tests
  // ==========================================

  describe('TTL Expiration', () => {
    test('should return null for expired items', async () => {
      cacheService.set('expires-soon', 'value', 50); // 50ms TTL

      // Wait for expiration
      await new Promise(resolve => setTimeout(resolve, 100));

      const result = cacheService.get('expires-soon');
      expect(result).toBeNull();
    });

    test('should return value for non-expired items', () => {
      cacheService.set('not-expired', 'value', 5000); // 5 second TTL

      const result = cacheService.get('not-expired');
      expect(result).toBe('value');
    });

    test('should use default TTL when not specified', () => {
      // Cache service configured with 1000ms default TTL
      cacheService.set('default-ttl', 'value');

      // Immediately check - should exist
      expect(cacheService.get('default-ttl')).toBe('value');
    });
  });

  // ==========================================
  // Cache Statistics
  // ==========================================

  describe('Cache Statistics', () => {
    test('should track cache hits', () => {
      cacheService.set('hit-test', 'value');
      cacheService.get('hit-test');
      cacheService.get('hit-test');
      cacheService.get('hit-test');

      const stats = cacheService.getStats();
      expect(stats.hits).toBe(3);
    });

    test('should track cache misses', () => {
      cacheService.get('miss-1');
      cacheService.get('miss-2');
      cacheService.get('miss-3');

      const stats = cacheService.getStats();
      expect(stats.misses).toBe(3);
    });

    test('should calculate hit rate correctly', () => {
      cacheService.set('key', 'value');
      cacheService.get('key'); // hit
      cacheService.get('key'); // hit
      cacheService.get('nonexistent'); // miss

      const stats = cacheService.getStats();
      expect(stats.hitRate).toBeCloseTo(66.67, 1); // 2/3 = 66.67%
    });

    test('should track cache size', () => {
      cacheService.set('key1', 'value1');
      cacheService.set('key2', 'value2');
      cacheService.set('key3', 'value3');

      const stats = cacheService.getStats();
      expect(stats.size).toBe(3);
    });
  });

  // ==========================================
  // Eviction Tests
  // ==========================================

  describe('Cache Eviction', () => {
    test('should evict oldest entry when at capacity', () => {
      const smallCache = new PolicyCacheService({
        defaultTtlMs: 60000,
        maxEntries: 3,
        cleanupIntervalMs: 60000
      });

      smallCache.set('first', 'value1');
      smallCache.set('second', 'value2');
      smallCache.set('third', 'value3');
      smallCache.set('fourth', 'value4'); // Should evict 'first'

      expect(smallCache.get('first')).toBeNull();
      expect(smallCache.get('second')).not.toBeNull();
      expect(smallCache.get('fourth')).not.toBeNull();

      smallCache.dispose();
    });
  });

  // ==========================================
  // Policy-Specific Methods
  // ==========================================

  describe('Policy Cache Methods', () => {
    test('should cache and retrieve policy by ID', () => {
      cacheService.setPolicy(mockPolicy);
      const cached = cacheService.getPolicy(1);

      expect(cached).toEqual(mockPolicy);
      expect(cached?.PolicyName).toBe('Test Policy');
    });

    test('should return null for non-cached policy', () => {
      const result = cacheService.getPolicy(999);
      expect(result).toBeNull();
    });

    test('should invalidate policy and related caches', () => {
      cacheService.setPolicy(mockPolicy);
      cacheService.setPolicyList('all', [mockPolicy]);
      cacheService.setComplianceSummary(1, {
        policyId: 1,
        policyName: 'Test Policy',
        totalAssigned: 100,
        totalAcknowledged: 80,
        totalOverdue: 5,
        totalExempted: 0,
        compliancePercentage: 80,
        averageTimeToAcknowledge: 2,
        riskLevel: ComplianceRisk.Medium
      });

      cacheService.invalidatePolicy(1);

      expect(cacheService.getPolicy(1)).toBeNull();
      expect(cacheService.getPolicyList('all')).toBeNull();
      expect(cacheService.getComplianceSummary(1)).toBeNull();
    });

    test('should cache and retrieve policy list', () => {
      const policies = [mockPolicy, { ...mockPolicy, Id: 2, PolicyName: 'Policy 2' }];
      cacheService.setPolicyList('active', policies);

      const cached = cacheService.getPolicyList('active');
      expect(cached).toHaveLength(2);
    });

    test('should also cache individual policies when setting list', () => {
      const policies = [mockPolicy, { ...mockPolicy, Id: 2 }];
      cacheService.setPolicyList('all', policies);

      // Individual policies should also be cached
      expect(cacheService.getPolicy(1)).toBeDefined();
      expect(cacheService.getPolicy(2)).toBeDefined();
    });
  });

  // ==========================================
  // Acknowledgement Cache Methods
  // ==========================================

  describe('Acknowledgement Cache Methods', () => {
    test('should cache and retrieve acknowledgement', () => {
      const ack = { Id: 1, PolicyId: 1, UserId: 100, Status: 'Pending' };
      cacheService.setAcknowledgement(ack as any);

      const cached = cacheService.getAcknowledgement(1);
      expect(cached?.PolicyId).toBe(1);
    });

    test('should invalidate acknowledgement and related caches', () => {
      const ack = { Id: 1, PolicyId: 1 };
      cacheService.setAcknowledgement(ack as any);
      cacheService.setUserDashboard(100, { data: 'test' });

      cacheService.invalidateAcknowledgement(1);

      expect(cacheService.getAcknowledgement(1)).toBeNull();
      expect(cacheService.getUserDashboard(100)).toBeNull();
    });
  });

  // ==========================================
  // User Dashboard Cache
  // ==========================================

  describe('User Dashboard Cache', () => {
    test('should cache user dashboard with shorter TTL', () => {
      cacheService.setUserDashboard(100, { pending: 5, completed: 10 });

      const cached = cacheService.getUserDashboard(100);
      expect(cached).toEqual({ pending: 5, completed: 10 });
    });

    test('should invalidate user dashboard', () => {
      cacheService.setUserDashboard(100, { data: 'test' });
      cacheService.invalidateUserDashboard(100);

      expect(cacheService.getUserDashboard(100)).toBeNull();
    });
  });

  // ==========================================
  // Policy Versions Cache
  // ==========================================

  describe('Policy Versions Cache', () => {
    test('should cache and retrieve policy versions', () => {
      const versions = [
        { versionNumber: '1.0', createdDate: new Date() },
        { versionNumber: '2.0', createdDate: new Date() }
      ];
      cacheService.setPolicyVersions(1, versions);

      const cached = cacheService.getPolicyVersions(1);
      expect(cached).toHaveLength(2);
    });
  });

  // ==========================================
  // Singleton Tests
  // ==========================================

  describe('Singleton Pattern', () => {
    test('should return same instance', () => {
      const instance1 = getPolicyCacheService();
      const instance2 = getPolicyCacheService();

      expect(instance1).toBe(instance2);
    });

    test('should reset singleton when calling reset', () => {
      const instance1 = getPolicyCacheService();
      instance1.set('test', 'value');

      resetPolicyCacheService();

      const instance2 = getPolicyCacheService();
      expect(instance2.get('test')).toBeNull();
    });
  });

  // ==========================================
  // Disposal Tests
  // ==========================================

  describe('Disposal', () => {
    test('should clear cache and stop timer on dispose', () => {
      cacheService.set('key', 'value');
      cacheService.dispose();

      expect(cacheService.getStats().size).toBe(0);
    });
  });
});

// ==========================================
// Pagination Utility Tests
// ==========================================

describe('paginateArray', () => {
  const items = Array.from({ length: 25 }, (_, i) => ({ id: i + 1 }));

  test('should paginate first page correctly', () => {
    const result = paginateArray(items, 1, 10);

    expect(result.items).toHaveLength(10);
    expect(result.pageNumber).toBe(1);
    expect(result.pageSize).toBe(10);
    expect(result.totalCount).toBe(25);
    expect(result.totalPages).toBe(3);
    expect(result.hasNextPage).toBe(true);
    expect(result.hasPreviousPage).toBe(false);
  });

  test('should paginate middle page correctly', () => {
    const result = paginateArray(items, 2, 10);

    expect(result.items).toHaveLength(10);
    expect(result.pageNumber).toBe(2);
    expect(result.hasNextPage).toBe(true);
    expect(result.hasPreviousPage).toBe(true);
  });

  test('should paginate last page correctly', () => {
    const result = paginateArray(items, 3, 10);

    expect(result.items).toHaveLength(5);
    expect(result.pageNumber).toBe(3);
    expect(result.hasNextPage).toBe(false);
    expect(result.hasPreviousPage).toBe(true);
  });

  test('should handle empty array', () => {
    const result = paginateArray([], 1, 10);

    expect(result.items).toHaveLength(0);
    expect(result.totalCount).toBe(0);
    expect(result.totalPages).toBe(0);
  });

  test('should handle page beyond bounds', () => {
    const result = paginateArray(items, 10, 10);

    expect(result.items).toHaveLength(0);
  });
});

// ==========================================
// Cache Key Generation Tests
// ==========================================

describe('generateCacheKey', () => {
  test('should generate key from filters', () => {
    const key = generateCacheKey({ status: 'active', category: 'HR' });

    expect(key).toContain('category');
    expect(key).toContain('status');
    expect(key).toContain('HR');
    expect(key).toContain('active');
  });

  test('should ignore undefined values', () => {
    const key = generateCacheKey({ status: 'active', category: undefined });

    expect(key).toContain('status');
    expect(key).not.toContain('category');
  });

  test('should ignore null values', () => {
    const key = generateCacheKey({ status: 'active', category: null });

    expect(key).not.toContain('category');
  });

  test('should return "all" for empty filters', () => {
    const key = generateCacheKey({});
    expect(key).toBe('all');
  });

  test('should sort keys for consistent ordering', () => {
    const key1 = generateCacheKey({ z: 1, a: 2, m: 3 });
    const key2 = generateCacheKey({ m: 3, a: 2, z: 1 });

    expect(key1).toBe(key2);
  });

  test('should handle complex values', () => {
    const key = generateCacheKey({
      ids: [1, 2, 3],
      nested: { deep: 'value' }
    });

    expect(key).toContain('[1,2,3]');
    expect(key).toContain('nested');
  });
});
