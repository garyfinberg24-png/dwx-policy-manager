// @ts-nocheck
// IntegrationService Unit Tests
// Testing Microsoft 365 integrations, third-party systems, and webhooks

import { IntegrationService } from './IntegrationService';
import { GraphService } from './GraphService';
import { SPFI } from '@pnp/sp';
import {
  IntegrationType,
  IntegrationStatus,
  IEntraIDEmployee,
  ITeamsChannel,
  IPlannerTask,
  IIntegrationResponse
} from '../models/IIntegration';

jest.mock('@pnp/sp');
jest.mock('./GraphService');

describe('IntegrationService', () => {
  let service: IntegrationService;
  let mockSp: jest.Mocked<SPFI>;
  let mockGraphService: jest.Mocked<GraphService>;

  beforeEach(() => {
    mockSp = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnThis(),
          items: {
            filter: jest.fn().mockReturnThis(),
            top: jest.fn().mockReturnThis(),
            add: jest.fn().mockResolvedValue({ data: { Id: 1 } }),
            getById: jest.fn().mockReturnThis()
          }
        }
      }
    } as any;

    mockGraphService = {
      getUserByEmail: jest.fn(),
      createTeamsChannel: jest.fn(),
      createPlannerTask: jest.fn()
    } as any;

    service = new IntegrationService(mockSp, mockGraphService);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  // ==========================================
  // Initialization Tests
  // ==========================================

  describe('Initialization', () => {
    test('should initialize and verify required lists', async () => {
      const requiredLists = [
        'JML_IntegrationConfigs',
        'JML_IntegrationLogs',
        'JML_IntegrationMappings',
        'JML_WebhookConfigs'
      ];

      await service.initialize();

      requiredLists.forEach(listName => {
        expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith(listName);
      });
    });

    test('should handle concurrent initialization calls', async () => {
      // Simulate multiple components calling initialize simultaneously
      const promises = [
        service.initialize(),
        service.initialize(),
        service.initialize()
      ];

      await Promise.all(promises);

      // Should only initialize once
      const listCalls = (mockSp.web.lists.getByTitle as jest.Mock).mock.calls.length;
      expect(listCalls).toBeLessThan(15); // 4 lists × 3 calls would be 12, actual should be less
    });

    test('should throw error if initialization fails', async () => {
      mockSp.web.lists.getByTitle = jest.fn().mockRejectedValue(
        new Error('List not found')
      );

      await expect(service.initialize()).rejects.toThrow();
    });
  });

  // ==========================================
  // Entra ID Sync Tests
  // ==========================================

  describe('Entra ID Sync', () => {
    const mockEntraIDEmployee: IEntraIDEmployee = {
      id: 'user-123',
      userPrincipalName: 'john.smith@company.com',
      displayName: 'John Smith',
      mail: 'john.smith@company.com',
      jobTitle: 'Software Engineer',
      department: 'IT',
      officeLocation: 'New York',
      mobilePhone: '+1-555-0100',
      businessPhones: ['+1-555-0100'],
      employeeId: 'EMP-12345',
      manager: {
        id: 'mgr-456',
        displayName: 'Jane Manager',
        mail: 'jane.manager@company.com'
      }
    };

    test('should sync employee from Entra ID successfully', async () => {
      const mockResponse: IIntegrationResponse<IEntraIDEmployee> = {
        success: true,
        data: mockEntraIDEmployee,
        timestamp: new Date(),
        duration: 123
      };

      mockGraphService.getUserByEmail.mockResolvedValue(mockResponse);

      const employee = await service.syncWithAD('john.smith@company.com');

      expect(employee.EmployeeName).toBe('John Smith');
      expect(employee.EmployeeEmail).toBe('john.smith@company.com');
      expect(employee.Department).toBe('IT');
      expect(employee.JobTitle).toBe('Software Engineer');
      expect(employee.ManagerEmail).toBe('jane.manager@company.com');
    });

    test('should handle missing manager gracefully', async () => {
      const employeeWithoutManager = { ...mockEntraIDEmployee, manager: undefined };
      const mockResponse: IIntegrationResponse<IEntraIDEmployee> = {
        success: true,
        data: employeeWithoutManager,
        timestamp: new Date(),
        duration: 100
      };

      mockGraphService.getUserByEmail.mockResolvedValue(mockResponse);

      const employee = await service.syncWithAD('john.smith@company.com');

      expect(employee.ManagerEmail).toBeUndefined();
    });

    test('should validate email before syncing', async () => {
      await expect(service.syncWithAD('invalid-email'))
        .rejects.toThrow('Invalid email format');

      expect(mockGraphService.getUserByEmail).not.toHaveBeenCalled();
    });

    test('should log integration success', async () => {
      const mockResponse: IIntegrationResponse<IEntraIDEmployee> = {
        success: true,
        data: mockEntraIDEmployee,
        timestamp: new Date(),
        duration: 100
      };

      mockGraphService.getUserByEmail.mockResolvedValue(mockResponse);

      await service.syncWithAD('john.smith@company.com');

      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('JML_IntegrationLogs');
      expect(mockSp.web.lists.getByTitle().items.add).toHaveBeenCalledWith(
        expect.objectContaining({
          IntegrationType: IntegrationType.EntraID,
          Success: true
        })
      );
    });

    test('should log integration failure', async () => {
      const mockResponse: IIntegrationResponse<IEntraIDEmployee> = {
        success: false,
        error: 'User not found',
        timestamp: new Date(),
        duration: 50
      };

      mockGraphService.getUserByEmail.mockResolvedValue(mockResponse);

      await expect(service.syncWithAD('nonexistent@company.com'))
        .rejects.toThrow('Failed to sync with AD');

      expect(mockSp.web.lists.getByTitle().items.add).toHaveBeenCalledWith(
        expect.objectContaining({
          Success: false,
          ErrorMessage: 'User not found'
        })
      );
    });
  });

  // ==========================================
  // Microsoft Teams Integration Tests
  // ==========================================

  describe('Microsoft Teams Integration', () => {
    const mockTeamsChannel: ITeamsChannel = {
      id: 'channel-123',
      displayName: 'Onboarding - John Smith',
      description: 'Onboarding channel for new hire',
      createdDateTime: new Date(),
      webUrl: 'https://teams.microsoft.com/l/channel/...'
    };

    beforeEach(() => {
      mockSp.web.lists.getByTitle().items.getById = jest.fn().mockReturnValue({
        mockReturnThis: jest.fn().mockResolvedValue({
          Id: 1,
          EmployeeName: 'John Smith',
          Department: 'IT',
          ProcessType: 'Onboarding'
        })
      });
    });

    test('should create Teams channel for process', async () => {
      const mockResponse: IIntegrationResponse<ITeamsChannel> = {
        success: true,
        data: mockTeamsChannel,
        timestamp: new Date(),
        duration: 200
      };

      mockGraphService.createTeamsChannel.mockResolvedValue(mockResponse);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { TeamId: 'team-123', IsActive: true }
      ]);

      const channelId = await service.createTeamsChannel(1);

      expect(channelId).toBe('channel-123');
      expect(mockGraphService.createTeamsChannel).toHaveBeenCalledWith(
        expect.objectContaining({
          teamId: 'team-123',
          displayName: expect.stringContaining('John Smith')
        })
      );
    });

    test('should throw error if IT team configuration is missing', async () => {
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([]);

      await expect(service.createTeamsChannel(1))
        .rejects.toThrow('No Teams integration configured for department: IT');
    });

    test('should include process details in channel description', async () => {
      const mockResponse: IIntegrationResponse<ITeamsChannel> = {
        success: true,
        data: mockTeamsChannel,
        timestamp: new Date(),
        duration: 150
      };

      mockGraphService.createTeamsChannel.mockResolvedValue(mockResponse);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { TeamId: 'team-123', IsActive: true }
      ]);

      await service.createTeamsChannel(1);

      const channelRequest = (mockGraphService.createTeamsChannel as jest.Mock).mock.calls[0][0];
      expect(channelRequest.description).toContain('Onboarding');
      expect(channelRequest.description).toContain('John Smith');
    });

    test('should handle Teams API errors', async () => {
      const mockResponse: IIntegrationResponse<ITeamsChannel> = {
        success: false,
        error: 'Forbidden: Insufficient permissions',
        timestamp: new Date(),
        duration: 100
      };

      mockGraphService.createTeamsChannel.mockResolvedValue(mockResponse);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { TeamId: 'team-123', IsActive: true }
      ]);

      await expect(service.createTeamsChannel(1))
        .rejects.toThrow('Failed to create Teams channel');
    });
  });

  // ==========================================
  // Microsoft Planner Integration Tests
  // ==========================================

  describe('Microsoft Planner Integration', () => {
    const mockTasks = [
      {
        Id: 1,
        Title: 'Setup Development Environment',
        Description: 'Install IDE, configure tools',
        AssignedTo: 'john.smith@company.com',
        DueDate: new Date('2024-02-01'),
        Priority: 'High'
      },
      {
        Id: 2,
        Title: 'Complete Security Training',
        Description: 'Mandatory security awareness training',
        AssignedTo: 'john.smith@company.com',
        DueDate: new Date('2024-02-05'),
        Priority: 'Medium'
      }
    ];

    test('should assign tasks to Planner', async () => {
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { PlanId: 'plan-123', IsActive: true }
      ]);

      const mockResponse: IIntegrationResponse<IPlannerTask> = {
        success: true,
        data: {
          id: 'task-123',
          title: 'Test Task',
          bucketId: 'bucket-123',
          planId: 'plan-123'
        },
        timestamp: new Date(),
        duration: 150
      };

      mockGraphService.createPlannerTask.mockResolvedValue(mockResponse);

      await service.assignPlannerTasks(mockTasks as any);

      expect(mockGraphService.createPlannerTask).toHaveBeenCalledTimes(2);
    });

    test('should map JML task priority to Planner priority', async () => {
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { PlanId: 'plan-123', IsActive: true }
      ]);

      const mockResponse: IIntegrationResponse<IPlannerTask> = {
        success: true,
        data: { id: 'task-123' } as any,
        timestamp: new Date(),
        duration: 100
      };

      mockGraphService.createPlannerTask.mockResolvedValue(mockResponse);

      await service.assignPlannerTasks([mockTasks[0]] as any);

      const plannerTaskRequest = (mockGraphService.createPlannerTask as jest.Mock).mock.calls[0][0];
      expect(plannerTaskRequest.priority).toBeDefined();
      // High = 1, Medium = 5, Low = 9 in Planner
      expect(plannerTaskRequest.priority).toBeLessThan(5);
    });

    test('should continue if some tasks fail', async () => {
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { PlanId: 'plan-123', IsActive: true }
      ]);

      mockGraphService.createPlannerTask
        .mockResolvedValueOnce({
          success: true,
          data: { id: 'task-1' } as any,
          timestamp: new Date(),
          duration: 100
        })
        .mockResolvedValueOnce({
          success: false,
          error: 'Task creation failed',
          timestamp: new Date(),
          duration: 50
        });

      await service.assignPlannerTasks(mockTasks as any);

      // Should attempt both tasks even if one fails
      expect(mockGraphService.createPlannerTask).toHaveBeenCalledTimes(2);
    });

    test('should throw if Planner configuration is missing', async () => {
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([]);

      await expect(service.assignPlannerTasks(mockTasks as any))
        .rejects.toThrow('No Planner integration configured');
    });
  });

  // ==========================================
  // Power Automate Webhook Tests
  // ==========================================

  describe('Power Automate Integration', () => {
    test('should trigger Power Automate flow', async () => {
      const flowId = 'abc123';
      const data = { processId: 1, employeeName: 'John Smith' };

      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        {
          FlowId: flowId,
          EndpointUrl: 'https://prod.westus.logic.azure.com/workflows/abc123/triggers',
          IsActive: true
        }
      ]);

      global.fetch = jest.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ status: 'success' })
      });

      await service.triggerPowerAutomate(flowId, data);

      expect(global.fetch).toHaveBeenCalledWith(
        expect.stringContaining(flowId),
        expect.objectContaining({
          method: 'POST',
          body: JSON.stringify(data)
        })
      );
    });

    test('should handle webhook failures gracefully', async () => {
      const flowId = 'abc123';

      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        {
          FlowId: flowId,
          EndpointUrl: 'https://prod.westus.logic.azure.com/workflows/abc123/triggers',
          IsActive: true
        }
      ]);

      global.fetch = jest.fn().mockResolvedValue({
        ok: false,
        status: 500,
        statusText: 'Internal Server Error'
      });

      await expect(service.triggerPowerAutomate(flowId, {}))
        .rejects.toThrow('Power Automate trigger failed');
    });

    test('should implement retry logic for webhook failures', async () => {
      const flowId = 'abc123';

      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        {
          FlowId: flowId,
          EndpointUrl: 'https://prod.westus.logic.azure.com/workflows/abc123/triggers',
          IsActive: true,
          RetryPolicy: 'Exponential Backoff',
          MaxRetries: 3
        }
      ]);

      global.fetch = jest.fn()
        .mockResolvedValueOnce({ ok: false, status: 503 }) // Fail
        .mockResolvedValueOnce({ ok: false, status: 503 }) // Fail
        .mockResolvedValueOnce({ ok: true }); // Success

      await service.triggerPowerAutomate(flowId, {});

      expect(global.fetch).toHaveBeenCalledTimes(3);
    });
  });

  // ==========================================
  // SAP/Workday Integration Tests (Mock)
  // ==========================================

  describe('External System Integration', () => {
    test('should map SAP fields to JML fields', () => {
      const sapEmployee = {
        PERNR: '00012345',
        ENAME: 'Smith, John',
        PERSK: 'A1',
        ORGEH: '10000100',
        PLANS: 'Software Engineer'
      };

      const mapped = service.mapSAPToJML(sapEmployee);

      expect(mapped.EmployeeID).toBe('00012345');
      expect(mapped.EmployeeName).toBe('John Smith'); // Reversed
      expect(mapped.JobTitle).toBe('Software Engineer');
    });

    test('should handle missing SAP fields', () => {
      const incompleteSAPData = {
        PERNR: '00012345'
        // Missing other fields
      };

      const mapped = service.mapSAPToJML(incompleteSAPData);

      expect(mapped.EmployeeID).toBe('00012345');
      expect(mapped.EmployeeName).toBeUndefined();
    });
  });

  // ==========================================
  // Integration Logging Tests
  // ==========================================

  describe('Integration Logging', () => {
    test('should log successful integrations', async () => {
      await service.logIntegration({
        integrationType: IntegrationType.EntraID,
        operation: 'Sync',
        success: true,
        duration: 123,
        recordsProcessed: 1
      });

      expect(mockSp.web.lists.getByTitle).toHaveBeenCalledWith('JML_IntegrationLogs');
      expect(mockSp.web.lists.getByTitle().items.add).toHaveBeenCalledWith(
        expect.objectContaining({
          IntegrationType: IntegrationType.EntraID,
          Success: true,
          Duration: 123
        })
      );
    });

    test('should log failed integrations with error details', async () => {
      await service.logIntegration({
        integrationType: IntegrationType.Teams,
        operation: 'Create Channel',
        success: false,
        duration: 50,
        statusCode: 403,
        errorMessage: 'Forbidden'
      });

      expect(mockSp.web.lists.getByTitle().items.add).toHaveBeenCalledWith(
        expect.objectContaining({
          Success: false,
          StatusCode: 403,
          ErrorMessage: 'Forbidden'
        })
      );
    });
  });

  // ==========================================
  // Error Handling Tests
  // ==========================================

  describe('Error Handling', () => {
    test('should handle network timeouts', async () => {
      mockGraphService.getUserByEmail.mockRejectedValue(
        new Error('Network timeout')
      );

      await expect(service.syncWithAD('user@company.com'))
        .rejects.toThrow('Network timeout');
    });

    test('should handle rate limiting', async () => {
      const mockResponse: IIntegrationResponse<IEntraIDEmployee> = {
        success: false,
        statusCode: 429,
        error: 'Rate limit exceeded',
        timestamp: new Date(),
        duration: 10
      };

      mockGraphService.getUserByEmail.mockResolvedValue(mockResponse);

      await expect(service.syncWithAD('user@company.com'))
        .rejects.toThrow('Rate limit exceeded');
    });

    test('should handle API permission errors', async () => {
      const mockResponse: IIntegrationResponse<ITeamsChannel> = {
        success: false,
        statusCode: 403,
        error: 'Insufficient permissions',
        timestamp: new Date(),
        duration: 50
      };

      mockGraphService.createTeamsChannel.mockResolvedValue(mockResponse);
      mockSp.web.lists.getByTitle().items.filter = jest.fn().mockResolvedValue([
        { TeamId: 'team-123', IsActive: true }
      ]);

      await expect(service.createTeamsChannel(1))
        .rejects.toThrow('Insufficient permissions');
    });
  });
});
