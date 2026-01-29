// @ts-nocheck
// Policy Service
// Comprehensive service for enterprise policy management
// Enhanced with server-side validation, audit logging, and retention management

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import {
  IPolicy,
  IPolicyVersion,
  IPolicyAcknowledgement,
  IPolicyExemption,
  IPolicyDistribution,
  IPolicyTemplate,
  IPolicyFeedback,
  IPolicyAuditLog,
  IPolicyPublishRequest,
  IPolicyAcknowledgeRequest,
  IPolicyComplianceSummary,
  IUserPolicyDashboard,
  IPolicyDashboardMetrics,
  PolicyStatus,
  AcknowledgementStatus,
  DistributionScope,
  ExemptionStatus,
  VersionType,
  DataClassification,
  RetentionCategory
} from '../models/IPolicy';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';
import { PolicyValidationService, PolicyOperation, IValidationResult, PolicyRole } from './PolicyValidationService';
import { PolicyAuditService, AuditEventType, AuditSeverity } from './PolicyAuditService';
import { PolicyRetentionService } from './PolicyRetentionService';
import {
  PolicyCacheService,
  getPolicyCacheService,
  IPaginatedResult,
  paginateArray,
  generateCacheKey
} from './PolicyCacheService';
import { PolicyNotificationService } from './PolicyNotificationService';
import { PolicyLists, QuizLists } from '../constants/SharePointListNames';

export class PolicyService {
  private sp: SPFI;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly POLICY_VERSIONS_LIST = PolicyLists.POLICY_VERSIONS;
  private readonly POLICY_ACKNOWLEDGEMENTS_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;
  private readonly POLICY_EXEMPTIONS_LIST = PolicyLists.POLICY_EXEMPTIONS;
  private readonly POLICY_DISTRIBUTIONS_LIST = PolicyLists.POLICY_DISTRIBUTIONS;
  private readonly POLICY_TEMPLATES_LIST = PolicyLists.POLICY_TEMPLATES;
  private readonly POLICY_FEEDBACK_LIST = PolicyLists.POLICY_FEEDBACK;
  private readonly POLICY_AUDIT_LOG_LIST = PolicyLists.POLICY_AUDIT_LOG;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserName: string = '';

  // Security services
  private validationService: PolicyValidationService;
  private auditService: PolicyAuditService;
  private retentionService: PolicyRetentionService;

  // Cache service
  private cacheService: PolicyCacheService;

  // Notification service
  private notificationService: PolicyNotificationService | null = null;
  private siteUrl: string = '';

  constructor(sp: SPFI, siteUrl?: string) {
    this.sp = sp;
    this.siteUrl = siteUrl || '';
    // Initialize security services
    this.validationService = new PolicyValidationService(sp);
    this.auditService = new PolicyAuditService(sp);
    this.retentionService = new PolicyRetentionService(sp);
    // Initialize cache service
    this.cacheService = getPolicyCacheService();
    // Initialize notification service
    if (this.siteUrl) {
      this.notificationService = new PolicyNotificationService(sp, this.siteUrl);
    }
  }

  /**
   * Initialize service with current user
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      this.currentUserName = user.Title;

      // Initialize audit service with current user context
      await this.auditService.initialize();
    } catch (error) {
      logger.error('PolicyService', 'Failed to initialize service:', error);
      throw error;
    }
  }

  /**
   * Get current user's role for policy operations
   */
  public async getCurrentUserRole(): Promise<PolicyRole> {
    // This would typically check user's SharePoint groups or a roles list
    // For now, return a default role - extend as needed
    try {
      const groups = await this.sp.web.currentUser.groups();
      const groupNames = groups.map(g => g.Title.toLowerCase());

      if (groupNames.some(g => g.includes('admin') || g.includes('owner'))) {
        return PolicyRole.Administrator;
      }
      if (groupNames.some(g => g.includes('compliance'))) {
        return PolicyRole.ComplianceOfficer;
      }
      if (groupNames.some(g => g.includes('publisher'))) {
        return PolicyRole.Publisher;
      }
      if (groupNames.some(g => g.includes('approver'))) {
        return PolicyRole.Approver;
      }
      if (groupNames.some(g => g.includes('reviewer'))) {
        return PolicyRole.Reviewer;
      }
      if (groupNames.some(g => g.includes('author'))) {
        return PolicyRole.Author;
      }
      return PolicyRole.Employee;
    } catch {
      return PolicyRole.Employee;
    }
  }

  // ============================================================================
  // POLICY CRUD OPERATIONS
  // ============================================================================

  /**
   * Create a new policy
   */
  public async createPolicy(policy: Partial<IPolicy>): Promise<IPolicy> {
    try {
      // Server-side validation: Check user permission
      const userRole = await this.getCurrentUserRole();
      const canCreate = await this.validationService.canPerformOperation(
        PolicyOperation.Create,
        undefined,
        { userRoles: [userRole], userId: this.currentUserId }
      );

      if (!canCreate.isValid) {
        // Log unauthorized attempt
        await this.auditService.logEvent({
          EventType: AuditEventType.UnauthorizedAccess,
          Severity: AuditSeverity.Security,
          EntityType: 'Policy',
          EntityId: 0,
          ActionDescription: `Unauthorized policy create attempt by user ${this.currentUserEmail}`,
          Metadata: JSON.stringify({ userRole, operation: PolicyOperation.Create })
        });
        throw new Error(canCreate.errors?.map(e => e.message).join(', ') || 'Unauthorized to create policies');
      }

      // Validate policy data
      const validation = await this.validationService.validatePolicyData(policy as IPolicy);
      if (!validation.isValid) {
        throw new Error(validation.errors?.map(e => e.message).join(', ') || 'Policy validation failed');
      }

      // Generate policy number if not provided
      if (!policy.PolicyNumber) {
        policy.PolicyNumber = await this.generatePolicyNumber(policy.PolicyCategory);
      }

      // Set defaults including data classification
      const policyData = {
        Title: policy.PolicyName,
        PolicyNumber: policy.PolicyNumber,
        PolicyName: policy.PolicyName,
        PolicyCategory: policy.PolicyCategory,
        PolicyType: policy.PolicyType,
        Description: policy.Description || '',
        VersionNumber: '0.1',
        VersionType: VersionType.Draft,
        MajorVersion: 0,
        MinorVersion: 1,
        PolicyStatus: PolicyStatus.Draft,
        PolicyOwnerId: policy.PolicyOwnerId || this.currentUserId,
        DocumentFormat: policy.DocumentFormat,
        RequiresAcknowledgement: policy.RequiresAcknowledgement ?? true,
        AcknowledgementType: policy.AcknowledgementType,
        RequiresQuiz: policy.RequiresQuiz ?? false,
        DistributionScope: policy.DistributionScope,
        ComplianceRisk: policy.ComplianceRisk,
        IsActive: false,
        IsMandatory: policy.IsMandatory ?? false,
        HTMLContent: policy.HTMLContent,
        DocumentURL: policy.DocumentURL,
        Tags: policy.Tags ? JSON.stringify(policy.Tags) : undefined,
        RelatedPolicyIds: policy.RelatedPolicyIds ? JSON.stringify(policy.RelatedPolicyIds) : undefined,
        // Data classification fields
        DataClassification: policy.DataClassification || DataClassification.Internal,
        RetentionCategory: policy.RetentionCategory || RetentionCategory.Standard,
        ContainsPII: policy.ContainsPII ?? false,
        ContainsPHI: policy.ContainsPHI ?? false,
        ContainsFinancialData: policy.ContainsFinancialData ?? false,
        ClassifiedById: this.currentUserId,
        ClassifiedDate: new Date().toISOString()
      };

      const result = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.add(policyData);

      // Enhanced audit logging using PolicyAuditService
      await this.auditService.logEvent({
        EventType: AuditEventType.PolicyCreated,
        Severity: AuditSeverity.Info,
        EntityType: 'Policy',
        EntityId: result.data.Id,
        EntityName: policy.PolicyName,
        PolicyId: result.data.Id,
        PolicyName: policy.PolicyName,
        ActionDescription: `Policy "${policy.PolicyName}" created with classification: ${policy.DataClassification || 'Internal'}`,
        ComplianceRelevant: true,
        DataClassification: policy.DataClassification || DataClassification.Internal
      });

      // Also log to legacy audit for backward compatibility
      await this.logAudit({
        EntityType: 'Policy',
        EntityId: result.data.Id,
        PolicyId: result.data.Id,
        Action: 'Created',
        ActionDescription: `Policy created: ${policy.PolicyName}`,
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return await this.getPolicyById(result.data.Id);
    } catch (error) {
      logger.error('PolicyService', 'Failed to create policy:', error);
      throw error;
    }
  }

  /**
   * Get policy by ID (with caching)
   */
  public async getPolicyById(policyId: number, bypassCache: boolean = false): Promise<IPolicy> {
    try {
      // Check cache first
      if (!bypassCache) {
        const cached = this.cacheService.getPolicy(policyId);
        if (cached) {
          return cached;
        }
      }

      const item = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .select(
          '*',
          'PolicyOwner/Id',
          'PolicyOwner/Title',
          'PolicyOwner/EMail'
        )
        .expand('PolicyOwner')();

      const policy = this.mapPolicyItem(item);

      // Cache the result
      this.cacheService.setPolicy(policy);

      return policy;
    } catch (error) {
      logger.error('PolicyService', 'Failed to get policy:', error);
      throw error;
    }
  }

  /**
   * Update policy
   */
  public async updatePolicy(policyId: number, updates: Partial<IPolicy>): Promise<IPolicy> {
    try {
      const updateData: Record<string, unknown> = { ...updates };

      // Handle array fields
      if (updates.Tags) {
        updateData.Tags = JSON.stringify(updates.Tags);
      }
      if (updates.RelatedPolicyIds) {
        updateData.RelatedPolicyIds = JSON.stringify(updates.RelatedPolicyIds);
      }

      await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .update(updateData);

      // Invalidate cache
      this.cacheService.invalidatePolicy(policyId);

      // Log audit
      await this.logAudit({
        EntityType: 'Policy',
        EntityId: policyId,
        PolicyId: policyId,
        Action: 'Updated',
        ActionDescription: `Policy updated: ${updates.PolicyName || 'Unknown'}`,
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return await this.getPolicyById(policyId, true); // Bypass cache to get fresh data
    } catch (error) {
      logger.error('PolicyService', 'Failed to update policy:', error);
      throw error;
    }
  }

  /**
   * Delete policy (soft delete by archiving)
   */
  public async deletePolicy(policyId: number): Promise<void> {
    try {
      await this.updatePolicy(policyId, {
        PolicyStatus: PolicyStatus.Archived,
        IsActive: false
      });

      // Log audit
      await this.logAudit({
        EntityType: 'Policy',
        EntityId: policyId,
        PolicyId: policyId,
        Action: 'Deleted',
        ActionDescription: 'Policy archived',
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });
    } catch (error) {
      logger.error('PolicyService', 'Failed to delete policy:', error);
      throw error;
    }
  }

  /**
   * Get all policies with optional filters (with caching)
   */
  public async getPolicies(filters?: {
    status?: PolicyStatus;
    category?: string;
    isActive?: boolean;
    isMandatory?: boolean;
  }, bypassCache: boolean = false): Promise<IPolicy[]> {
    try {
      // Generate cache key from filters
      const cacheKey = generateCacheKey(filters || {});

      // Check cache first
      if (!bypassCache) {
        const cached = this.cacheService.getPolicyList(cacheKey);
        if (cached) {
          return cached;
        }
      }

      let query = this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.select(
          '*',
          'PolicyOwner/Id',
          'PolicyOwner/Title',
          'PolicyOwner/EMail'
        )
        .expand('PolicyOwner');

      // Apply filters
      const filterConditions: string[] = [];
      if (filters?.status) {
        filterConditions.push(`PolicyStatus eq '${filters.status}'`);
      }
      if (filters?.category) {
        filterConditions.push(`PolicyCategory eq '${filters.category}'`);
      }
      if (filters?.isActive !== undefined) {
        filterConditions.push(`IsActive eq ${filters.isActive}`);
      }
      if (filters?.isMandatory !== undefined) {
        filterConditions.push(`IsMandatory eq ${filters.isMandatory}`);
      }

      if (filterConditions.length > 0) {
        query = query.filter(filterConditions.join(' and '));
      }

      const items = await query.top(5000)();
      const policies = items.map(item => this.mapPolicyItem(item));

      // Cache the results
      this.cacheService.setPolicyList(cacheKey, policies);

      return policies;
    } catch (error) {
      logger.error('PolicyService', 'Failed to get policies:', error);
      throw error;
    }
  }

  /**
   * Get policies with pagination support
   */
  public async getPoliciesPaginated(
    pageNumber: number = 1,
    pageSize: number = 20,
    filters?: {
      status?: PolicyStatus;
      category?: string;
      isActive?: boolean;
      isMandatory?: boolean;
      searchTerm?: string;
      sortBy?: string;
      sortDirection?: 'asc' | 'desc';
    },
    bypassCache: boolean = false
  ): Promise<IPaginatedResult<IPolicy>> {
    try {
      // Get all matching policies (with caching)
      const allPolicies = await this.getPolicies({
        status: filters?.status,
        category: filters?.category,
        isActive: filters?.isActive,
        isMandatory: filters?.isMandatory
      }, bypassCache);

      // Apply search filter if provided
      let filteredPolicies = allPolicies;
      if (filters?.searchTerm) {
        const searchLower = filters.searchTerm.toLowerCase();
        filteredPolicies = allPolicies.filter(policy =>
          policy.PolicyName?.toLowerCase().includes(searchLower) ||
          policy.PolicyNumber?.toLowerCase().includes(searchLower) ||
          policy.Description?.toLowerCase().includes(searchLower) ||
          policy.Tags?.some(tag => tag.toLowerCase().includes(searchLower))
        );
      }

      // Apply sorting
      if (filters?.sortBy) {
        const sortField = filters.sortBy;
        filteredPolicies.sort((a, b) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const aValue = (a as any)[sortField];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const bValue = (b as any)[sortField];

          if (aValue === undefined || aValue === null) return 1;
          if (bValue === undefined || bValue === null) return -1;

          let comparison = 0;
          if (typeof aValue === 'string' && typeof bValue === 'string') {
            comparison = aValue.localeCompare(bValue);
          } else if (aValue < bValue) {
            comparison = -1;
          } else if (aValue > bValue) {
            comparison = 1;
          }

          return filters.sortDirection === 'desc' ? -comparison : comparison;
        });
      }

      // Apply pagination
      return paginateArray(filteredPolicies, pageNumber, pageSize);
    } catch (error) {
      logger.error('PolicyService', 'Failed to get paginated policies:', error);
      throw error;
    }
  }

  /**
   * Get cache statistics
   */
  public getCacheStats(): { hits: number; misses: number; size: number; hitRate: number } {
    return this.cacheService.getStats();
  }

  /**
   * Clear policy cache
   */
  public clearCache(): void {
    this.cacheService.clear();
  }

  /**
   * Refresh cache for a specific policy
   */
  public async refreshPolicyCache(policyId: number): Promise<IPolicy> {
    this.cacheService.invalidatePolicy(policyId);
    return this.getPolicyById(policyId, true);
  }

  // ============================================================================
  // POLICY LIFECYCLE MANAGEMENT
  // ============================================================================

  /**
   * Submit policy for review
   */
  public async submitForReview(policyId: number, reviewerIds: number[]): Promise<IPolicy> {
    try {
      await this.updatePolicy(policyId, {
        PolicyStatus: PolicyStatus.InReview,
        ReviewerIds: reviewerIds,
        SubmittedForReviewDate: new Date()
      });

      // TODO: Send notifications to reviewers

      await this.logAudit({
        EntityType: 'Policy',
        EntityId: policyId,
        PolicyId: policyId,
        Action: 'SubmittedForReview',
        ActionDescription: 'Policy submitted for review',
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return await this.getPolicyById(policyId);
    } catch (error) {
      logger.error('PolicyService', 'Failed to submit for review:', error);
      throw error;
    }
  }

  /**
   * Get all policies (alias for getPolicies without filters)
   */
  public async getAllPolicies(): Promise<IPolicy[]> {
    return this.getPolicies();
  }

  /**
   * Reject policy
   */
  public async rejectPolicy(policyId: number, rejectionReason: string): Promise<IPolicy> {
    try {
      // Server-side validation: Check user has rejection permission
      const userRole = await this.getCurrentUserRole();
      const canReject = await this.validationService.canPerformOperation(
        PolicyOperation.Reject,
        policyId,
        { userRoles: [userRole], userId: this.currentUserId }
      );

      if (!canReject.isValid) {
        // Log unauthorized rejection attempt
        await this.auditService.logEvent({
          EventType: AuditEventType.UnauthorizedAccess,
          Severity: AuditSeverity.Security,
          EntityType: 'Policy',
          EntityId: policyId,
          PolicyId: policyId,
          ActionDescription: `Unauthorized policy rejection attempt by user ${this.currentUserEmail}`,
          Metadata: JSON.stringify({ userRole, operation: PolicyOperation.Reject })
        });
        throw new Error(canReject.errors?.map(e => e.message).join(', ') || 'Unauthorized to reject policies');
      }

      // Validate rejection reason is provided
      if (!rejectionReason || rejectionReason.trim().length === 0) {
        throw new Error('Rejection reason is required');
      }

      // Get current policy for audit
      const currentPolicy = await this.getPolicyById(policyId);

      await this.updatePolicy(policyId, {
        PolicyStatus: PolicyStatus.Draft, // Return to draft for revision
        RejectionReason: rejectionReason,
        RejectedDate: new Date()
      });

      // Enhanced audit logging - Policy Rejection
      await this.auditService.logPolicyRejection(currentPolicy, rejectionReason);

      // Legacy audit log for backward compatibility
      await this.logAudit({
        EntityType: 'Policy',
        EntityId: policyId,
        PolicyId: policyId,
        Action: 'Rejected',
        ActionDescription: rejectionReason,
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return await this.getPolicyById(policyId);
    } catch (error) {
      logger.error('PolicyService', 'Failed to reject policy:', error);
      throw error;
    }
  }

  /**
   * Archive policy
   */
  public async archivePolicy(policyId: number, archiveReason?: string): Promise<IPolicy> {
    try {
      await this.updatePolicy(policyId, {
        PolicyStatus: PolicyStatus.Archived,
        IsActive: false,
        ArchivedDate: new Date()
      });

      await this.logAudit({
        EntityType: 'Policy',
        EntityId: policyId,
        PolicyId: policyId,
        Action: 'Archived',
        ActionDescription: archiveReason || 'Policy archived',
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return await this.getPolicyById(policyId);
    } catch (error) {
      logger.error('PolicyService', 'Failed to archive policy:', error);
      throw error;
    }
  }

  /**
   * Approve policy
   */
  public async approvePolicy(policyId: number, approverComments?: string): Promise<IPolicy> {
    try {
      // Server-side validation: Check user has approval permission
      const userRole = await this.getCurrentUserRole();
      const canApprove = await this.validationService.canPerformOperation(
        PolicyOperation.Approve,
        policyId,
        { userRoles: [userRole], userId: this.currentUserId }
      );

      if (!canApprove.isValid) {
        // Log unauthorized approval attempt
        await this.auditService.logEvent({
          EventType: AuditEventType.UnauthorizedAccess,
          Severity: AuditSeverity.Security,
          EntityType: 'Policy',
          EntityId: policyId,
          PolicyId: policyId,
          ActionDescription: `Unauthorized policy approval attempt by user ${this.currentUserEmail}`,
          Metadata: JSON.stringify({ userRole, operation: PolicyOperation.Approve })
        });
        throw new Error(canApprove.errors?.map(e => e.message).join(', ') || 'Unauthorized to approve policies');
      }

      // Get current policy to validate status transition
      const currentPolicy = await this.getPolicyById(policyId);
      if (currentPolicy.PolicyStatus !== PolicyStatus.InReview) {
        throw new Error(`Cannot approve policy in ${currentPolicy.PolicyStatus} status. Policy must be In Review.`);
      }

      await this.updatePolicy(policyId, {
        PolicyStatus: PolicyStatus.Approved,
        ApprovedDate: new Date()
      });

      // Enhanced audit logging - Policy Approval
      await this.auditService.logPolicyApproval(currentPolicy, approverComments);

      // Legacy audit log for backward compatibility
      await this.logAudit({
        EntityType: 'Policy',
        EntityId: policyId,
        PolicyId: policyId,
        Action: 'Approved',
        ActionDescription: approverComments || 'Policy approved',
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return await this.getPolicyById(policyId);
    } catch (error) {
      logger.error('PolicyService', 'Failed to approve policy:', error);
      throw error;
    }
  }

  /**
   * Publish policy and distribute to users
   */
  public async publishPolicy(request: IPolicyPublishRequest): Promise<IPolicyDistribution> {
    try {
      const policy = await this.getPolicyById(request.policyId);

      // Update policy status
      await this.updatePolicy(request.policyId, {
        PolicyStatus: PolicyStatus.Published,
        IsActive: true,
        PublishedDate: new Date(),
        EffectiveDate: request.effectiveDate || new Date()
      });

      // Create new version
      await this.createVersion(request.policyId, VersionType.Major, 'Policy published');

      // Get target users
      const targetUsers = await this.resolveTargetUsers(
        request.distributionScope,
        request.targetUserIds,
        request.targetDepartments,
        request.targetLocations,
        request.targetRoles
      );

      // Create distribution record
      const distribution = await this.createDistribution({
        PolicyId: request.policyId,
        DistributionName: `${policy.PolicyName} - ${new Date().toLocaleDateString()}`,
        DistributionScope: request.distributionScope,
        ScheduledDate: new Date(),
        TargetCount: targetUsers.length,
        DueDate: request.dueDate,
        IsActive: true
      });

      // Create acknowledgement records for each user
      await this.createAcknowledgements(
        request.policyId,
        policy.VersionNumber,
        targetUsers,
        request.dueDate
      );

      // Send notifications
      if (request.sendNotifications && this.notificationService) {
        const policy = await this.getPolicyById(request.policyId);
        await this.notificationService.sendNewPolicyNotification(policy, targetUsers);
        logger.info('PolicyService', `Sent notifications to ${targetUsers.length} users for policy ${policy.PolicyName}`);
      }

      await this.logAudit({
        EntityType: 'Policy',
        EntityId: request.policyId,
        PolicyId: request.policyId,
        Action: 'Published',
        ActionDescription: `Policy published to ${targetUsers.length} users`,
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return distribution;
    } catch (error) {
      logger.error('PolicyService', 'Failed to publish policy:', error);
      throw error;
    }
  }

  // ============================================================================
  // VERSION MANAGEMENT
  // ============================================================================

  /**
   * Create a new version of a policy
   */
  private async createVersion(
    policyId: number,
    versionType: VersionType,
    changeDescription: string
  ): Promise<IPolicyVersion> {
    try {
      const policy = await this.getPolicyById(policyId);

      // Mark all previous versions as not current
      const versions = await this.getPolicyVersions(policyId);
      for (const version of versions) {
        if (version.IsCurrentVersion) {
          await this.sp.web.lists
            .getByTitle(this.POLICY_VERSIONS_LIST)
            .items.getById(version.Id!)
            .update({ IsCurrentVersion: false });
        }
      }

      // Create new version
      const versionData = {
        Title: `${policy.PolicyName} - v${policy.VersionNumber}`,
        PolicyId: policyId,
        VersionNumber: policy.VersionNumber,
        VersionType: versionType,
        ChangeDescription: changeDescription,
        DocumentURL: policy.DocumentURL,
        HTMLContent: policy.HTMLContent,
        EffectiveDate: new Date().toISOString(),
        CreatedById: this.currentUserId,
        IsCurrentVersion: true
      };

      const result = await this.sp.web.lists
        .getByTitle(this.POLICY_VERSIONS_LIST)
        .items.add(versionData);

      return result.data as IPolicyVersion;
    } catch (error) {
      logger.error('PolicyService', 'Failed to create version:', error);
      throw error;
    }
  }

  /**
   * Get all versions of a policy
   */
  public async getPolicyVersions(policyId: number): Promise<IPolicyVersion[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.POLICY_VERSIONS_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .orderBy('Created', false)
        .top(100)();

      return items as IPolicyVersion[];
    } catch (error) {
      logger.error('PolicyService', 'Failed to get policy versions:', error);
      throw error;
    }
  }

  // ============================================================================
  // ACKNOWLEDGEMENT MANAGEMENT
  // ============================================================================

  /**
   * Create acknowledgement records for users
   */
  private async createAcknowledgements(
    policyId: number,
    versionNumber: string,
    userIds: number[],
    dueDate?: Date
  ): Promise<void> {
    try {
      const policy = await this.getPolicyById(policyId);
      const batchSize = 100;

      for (let i = 0; i < userIds.length; i += batchSize) {
        const batch = userIds.slice(i, i + batchSize);
        const promises = batch.map(async (userId) => {
          // Check if acknowledgement already exists
          const existing = await this.sp.web.lists
            .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
            .items.filter(`PolicyId eq ${policyId} and AckUserId eq ${userId} and PolicyVersionNumber eq '${versionNumber}'`)
            .top(1)();

          if (existing.length === 0) {
            const ackData = {
              Title: `${policy.PolicyName} - User ${userId}`,
              PolicyId: policyId,
              PolicyVersionNumber: versionNumber,
              AckUserId: userId,
              UserEmail: '', // TODO: Fetch from user profile
              AckStatus: AcknowledgementStatus.Sent,
              AssignedDate: new Date().toISOString(),
              DueDate: dueDate?.toISOString(),
              QuizRequired: policy.RequiresQuiz,
              DocumentOpenCount: 0,
              TotalReadTimeSeconds: 0,
              IsDelegated: false,
              RemindersSent: 0,
              IsExempted: false,
              IsCompliant: false
            };

            await this.sp.web.lists
              .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
              .items.add(ackData);
          }
        });

        await Promise.all(promises);
      }
    } catch (error) {
      logger.error('PolicyService', 'Failed to create acknowledgements:', error);
      throw error;
    }
  }

  /**
   * Get user's acknowledgement for a policy
   */
  public async getUserAcknowledgement(policyId: number, userId: number): Promise<IPolicyAcknowledgement | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`PolicyId eq ${policyId} and AckUserId eq ${userId}`)
        .orderBy('Created', false)
        .top(1)();

      return items.length > 0 ? (items[0] as IPolicyAcknowledgement) : null;
    } catch (error) {
      logger.error('PolicyService', 'Failed to get user acknowledgement:', error);
      throw error;
    }
  }

  /**
   * Acknowledge a policy
   */
  public async acknowledgePolicy(request: IPolicyAcknowledgeRequest): Promise<IPolicyAcknowledgement> {
    try {
      // Get the acknowledgement record first
      const existingAck = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(request.acknowledgementId)() as IPolicyAcknowledgement;

      // Server-side validation: Check user can acknowledge
      const userRole = await this.getCurrentUserRole();
      const canAcknowledge = await this.validationService.canPerformOperation(
        PolicyOperation.Acknowledge,
        existingAck.PolicyId,
        { userRoles: [userRole], userId: this.currentUserId }
      );

      if (!canAcknowledge.isValid) {
        await this.auditService.logEvent({
          EventType: AuditEventType.UnauthorizedAccess,
          Severity: AuditSeverity.Warning,
          EntityType: 'Acknowledgement',
          EntityId: request.acknowledgementId,
          PolicyId: existingAck.PolicyId,
          ActionDescription: `Unauthorized acknowledgement attempt by user ${this.currentUserEmail}`,
          Metadata: JSON.stringify({ acknowledgementId: request.acknowledgementId })
        });
        throw new Error(canAcknowledge.errors?.map(e => e.message).join(', ') || 'Unauthorized to acknowledge this policy');
      }

      // Validate acknowledgement request
      const validation = await this.validationService.validateAcknowledgement(request.acknowledgementId, this.currentUserId);
      if (!validation.isValid) {
        throw new Error(validation.errors?.map(e => e.message).join(', ') || 'Acknowledgement validation failed');
      }

      // Verify acknowledgement is for the current user
      if (existingAck.AckUserId !== this.currentUserId && !existingAck.IsDelegated) {
        throw new Error('Cannot acknowledge policy on behalf of another user without delegation');
      }

      // Get the policy to check retention category
      const policy = await this.getPolicyById(existingAck.PolicyId);

      // Calculate retention expiry based on policy retention category
      const retentionDays = this.retentionService.getRetentionPeriodDays(
        policy.RetentionCategory || RetentionCategory.Standard
      );
      const retentionExpiryDate = new Date();
      retentionExpiryDate.setDate(retentionExpiryDate.getDate() + retentionDays);

      const updateData = {
        AckStatus: AcknowledgementStatus.Acknowledged,
        AcknowledgedDate: request.acknowledgedDate.toISOString(),
        DigitalSignature: request.digitalSignature,
        IsCompliant: true,
        ComplianceDate: new Date().toISOString(),
        // Apply retention policy to acknowledgement record
        RetentionExpiryDate: retentionExpiryDate.toISOString(),
        RetentionCategory: policy.RetentionCategory || RetentionCategory.Standard,
        DataClassification: policy.DataClassification || DataClassification.Internal
      };

      await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(request.acknowledgementId)
        .update(updateData);

      // Get the updated acknowledgement
      const ack = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(request.acknowledgementId)() as IPolicyAcknowledgement;

      // Enhanced audit logging - Acknowledgement completion
      await this.auditService.logAcknowledgement(ack, policy, 'Digital Signature');

      // Legacy audit log for backward compatibility
      await this.logAudit({
        EntityType: 'Acknowledgement',
        EntityId: request.acknowledgementId,
        PolicyId: ack.PolicyId,
        Action: 'Acknowledged',
        ActionDescription: `Policy acknowledged by user. Retention: ${policy.RetentionCategory}, Classification: ${policy.DataClassification}`,
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return ack;
    } catch (error) {
      logger.error('PolicyService', 'Failed to acknowledge policy:', error);
      throw error;
    }
  }

  /**
   * Track policy document opening
   */
  public async trackPolicyOpen(acknowledgementId: number): Promise<void> {
    try {
      const ack = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(acknowledgementId)() as IPolicyAcknowledgement;

      const updateData: any = {
        DocumentOpenCount: (ack.DocumentOpenCount || 0) + 1,
        LastAccessedDate: new Date().toISOString()
      };

      if (!ack.FirstOpenedDate) {
        updateData.FirstOpenedDate = new Date().toISOString();
        updateData.AckStatus = AcknowledgementStatus.Opened;
      }

      await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(acknowledgementId)
        .update(updateData);
    } catch (error) {
      logger.error('PolicyService', 'Failed to track policy open:', error);
      throw error;
    }
  }

  // ============================================================================
  // EXEMPTIONS
  // ============================================================================

  /**
   * Request policy exemption
   */
  public async requestExemption(exemption: Partial<IPolicyExemption>): Promise<IPolicyExemption> {
    try {
      const exemptionData = {
        Title: `Exemption - Policy ${exemption.PolicyId} - User ${exemption.UserId}`,
        PolicyId: exemption.PolicyId,
        UserId: exemption.UserId,
        ExemptionReason: exemption.ExemptionReason,
        ExemptionType: exemption.ExemptionType,
        Status: ExemptionStatus.Pending,
        RequestDate: new Date().toISOString(),
        RequestedById: this.currentUserId
      };

      const result = await this.sp.web.lists
        .getByTitle(this.POLICY_EXEMPTIONS_LIST)
        .items.add(exemptionData);

      await this.logAudit({
        EntityType: 'Exemption',
        EntityId: result.data.Id,
        PolicyId: exemption.PolicyId,
        Action: 'ExemptionRequested',
        ActionDescription: `Exemption requested: ${exemption.ExemptionReason}`,
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return result.data as IPolicyExemption;
    } catch (error) {
      logger.error('PolicyService', 'Failed to request exemption:', error);
      throw error;
    }
  }

  /**
   * Approve exemption
   */
  public async approveExemption(exemptionId: number, comments?: string): Promise<IPolicyExemption> {
    try {
      const updateData = {
        Status: ExemptionStatus.Approved,
        ApprovedById: this.currentUserId,
        ApprovedDate: new Date().toISOString(),
        ReviewComments: comments
      };

      await this.sp.web.lists
        .getByTitle(this.POLICY_EXEMPTIONS_LIST)
        .items.getById(exemptionId)
        .update(updateData);

      const exemption = await this.sp.web.lists
        .getByTitle(this.POLICY_EXEMPTIONS_LIST)
        .items.getById(exemptionId)() as IPolicyExemption;

      // Update acknowledgement
      const ack = await this.getUserAcknowledgement(exemption.PolicyId, exemption.UserId!);
      if (ack) {
        await this.sp.web.lists
          .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
          .items.getById(ack.Id!)
          .update({
            IsExempted: true,
            ExemptionId: exemptionId,
            AckStatus: AcknowledgementStatus.Exempted
          });
      }

      await this.logAudit({
        EntityType: 'Exemption',
        EntityId: exemptionId,
        PolicyId: exemption.PolicyId,
        Action: 'ExemptionApproved',
        ActionDescription: comments || 'Exemption approved',
        PerformedById: this.currentUserId,
        PerformedByEmail: this.currentUserEmail,
        ActionDate: new Date(),
        ComplianceRelevant: true
      });

      return exemption;
    } catch (error) {
      logger.error('PolicyService', 'Failed to approve exemption:', error);
      throw error;
    }
  }

  // ============================================================================
  // DASHBOARDS & REPORTING
  // ============================================================================

  /**
   * Get user's policy dashboard
   */
  public async getUserDashboard(userId: number): Promise<IUserPolicyDashboard> {
    try {
      const allAcknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`AckUserId eq ${userId}`)
        .select('*')
        .top(1000)();

      const acks = allAcknowledgements as IPolicyAcknowledgement[];

      const pending = acks.filter(a =>
        a.AckStatus === AcknowledgementStatus.Sent ||
        a.AckStatus === AcknowledgementStatus.Opened
      );
      const overdue = acks.filter(a => a.AckStatus === AcknowledgementStatus.Overdue);
      const completed = acks.filter(a => a.AckStatus === AcknowledgementStatus.Acknowledged);

      return {
        userId,
        pendingAcknowledgements: pending,
        overdueAcknowledgements: overdue,
        completedAcknowledgements: completed,
        totalPending: pending.length,
        totalOverdue: overdue.length,
        totalCompleted: completed.length,
        complianceScore: acks.length > 0 ? (completed.length / acks.length) * 100 : 100
      };
    } catch (error) {
      logger.error('PolicyService', 'Failed to get user dashboard:', error);
      throw error;
    }
  }

  /**
   * Get policy compliance summary
   */
  public async getPolicyComplianceSummary(policyId: number): Promise<IPolicyComplianceSummary> {
    try {
      const policy = await this.getPolicyById(policyId);
      const acknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .top(5000)();

      const acks = acknowledgements as IPolicyAcknowledgement[];
      const acknowledged = acks.filter(a => a.AckStatus === AcknowledgementStatus.Acknowledged);
      const overdue = acks.filter(a => a.AckStatus === AcknowledgementStatus.Overdue);
      const exempted = acks.filter(a => a.AckStatus === AcknowledgementStatus.Exempted);

      // Calculate average time to acknowledge
      const acknowledgedTimes = acknowledged
        .filter(a => a.AcknowledgedDate && a.AssignedDate)
        .map(a => {
          const assigned = new Date(a.AssignedDate!);
          const acked = new Date(a.AcknowledgedDate!);
          return (acked.getTime() - assigned.getTime()) / (1000 * 60 * 60 * 24); // days
        });

      const avgTime = acknowledgedTimes.length > 0
        ? acknowledgedTimes.reduce((sum, time) => sum + time, 0) / acknowledgedTimes.length
        : 0;

      return {
        policyId,
        policyName: policy.PolicyName,
        totalAssigned: acks.length,
        totalAcknowledged: acknowledged.length,
        totalOverdue: overdue.length,
        totalExempted: exempted.length,
        compliancePercentage: acks.length > 0 ? (acknowledged.length / acks.length) * 100 : 0,
        averageTimeToAcknowledge: avgTime,
        riskLevel: policy.ComplianceRisk
      };
    } catch (error) {
      logger.error('PolicyService', 'Failed to get compliance summary:', error);
      throw error;
    }
  }

  /**
   * Get overall policy dashboard metrics
   */
  public async getDashboardMetrics(): Promise<IPolicyDashboardMetrics> {
    try {
      const policies = await this.getPolicies();
      const activePolicies = policies.filter(p => p.IsActive);
      const draftPolicies = policies.filter(p => p.PolicyStatus === PolicyStatus.Draft);

      // Policies expiring in next 30 days
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);
      const expiringSoon = policies.filter(p =>
        p.ExpiryDate && new Date(p.ExpiryDate) <= thirtyDaysFromNow
      );

      const allAcknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.top(5000)();

      const acks = allAcknowledgements as IPolicyAcknowledgement[];
      const acknowledged = acks.filter(a => a.AckStatus === AcknowledgementStatus.Acknowledged);
      const overdue = acks.filter(a => a.AckStatus === AcknowledgementStatus.Overdue);

      const criticalRisk = policies.filter(p => p.ComplianceRisk === 'Critical');

      // Get recent feedback count (last 30 days)
      let recentFeedbackCount = 0;
      try {
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        const recentFeedback = await this.sp.web.lists
          .getByTitle(this.POLICY_FEEDBACK_LIST)
          .items
          .filter(`Created ge datetime'${thirtyDaysAgo.toISOString()}'`)
          .select('Id')();
        recentFeedbackCount = recentFeedback.length;
      } catch {
        // Feedback list may not exist yet
        recentFeedbackCount = 0;
      }

      return {
        totalPolicies: policies.length,
        activePolicies: activePolicies.length,
        draftPolicies: draftPolicies.length,
        expiringSoon: expiringSoon.length,
        overallComplianceRate: acks.length > 0 ? (acknowledged.length / acks.length) * 100 : 100,
        totalAcknowledgements: acks.length,
        overdueAcknowledgements: overdue.length,
        criticalRiskPolicies: criticalRisk.length,
        recentFeedback: recentFeedbackCount
      };
    } catch (error) {
      logger.error('PolicyService', 'Failed to get dashboard metrics:', error);
      throw error;
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Generate unique policy number
   */
  private async generatePolicyNumber(category?: string): Promise<string> {
    try {
      const prefix = category ? category.substring(0, 3).toUpperCase() : 'POL';
      const policies = await this.getPolicies();
      const count = policies.length + 1;
      return `${prefix}-${String(count).padStart(4, '0')}`;
    } catch (error) {
      return `POL-${String(Date.now()).slice(-4)}`;
    }
  }

  /**
   * Resolve target users based on distribution scope
   */
  private async resolveTargetUsers(
    scope: DistributionScope,
    userIds?: number[],
    departments?: string[],
    locations?: string[],
    roles?: string[]
  ): Promise<number[]> {
    try {
      let targetUsers: number[] = [];

      switch (scope) {
        case DistributionScope.AllEmployees:
          // Get all site users
          const allUsers = await this.sp.web.siteUsers();
          targetUsers = allUsers.map(u => u.Id);
          break;

        case DistributionScope.Custom:
          targetUsers = userIds || [];
          break;

        case DistributionScope.Department:
          // Filter users by department from PM_Employees list
          if (departments && departments.length > 0) {
            try {
              const deptFilter = departments.map(d => `Department eq '${d}'`).join(' or ');
              const deptEmployees = await this.sp.web.lists
                .getByTitle('PM_Employees')
                .items
                .filter(deptFilter)
                .select('UserId')();
              targetUsers = deptEmployees.map((e: Record<string, unknown>) => e.UserId as number).filter(Boolean);
            } catch {
              targetUsers = userIds || [];
            }
          } else {
            targetUsers = userIds || [];
          }
          break;

        case DistributionScope.Location:
          // Filter users by location from PM_Employees list
          if (locations && locations.length > 0) {
            try {
              const locFilter = locations.map(l => `Location eq '${l}'`).join(' or ');
              const locEmployees = await this.sp.web.lists
                .getByTitle('PM_Employees')
                .items
                .filter(locFilter)
                .select('UserId')();
              targetUsers = locEmployees.map((e: Record<string, unknown>) => e.UserId as number).filter(Boolean);
            } catch {
              targetUsers = userIds || [];
            }
          } else {
            targetUsers = userIds || [];
          }
          break;

        case DistributionScope.Role:
          // Filter users by role/job title from PM_Employees list
          if (roles && roles.length > 0) {
            try {
              const roleFilter = roles.map(r => `JobTitle eq '${r}'`).join(' or ');
              const roleEmployees = await this.sp.web.lists
                .getByTitle('PM_Employees')
                .items
                .filter(roleFilter)
                .select('UserId')();
              targetUsers = roleEmployees.map((e: Record<string, unknown>) => e.UserId as number).filter(Boolean);
            } catch {
              targetUsers = userIds || [];
            }
          } else {
            targetUsers = userIds || [];
          }
          break;

        case DistributionScope.NewHiresOnly:
          // Get new hires from PM_Employees (hired in last 90 days)
          try {
            const ninetyDaysAgo = new Date();
            ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
            const newHires = await this.sp.web.lists
              .getByTitle('PM_Employees')
              .items
              .filter(`HireDate ge datetime'${ninetyDaysAgo.toISOString()}'`)
              .select('UserId')();
            targetUsers = newHires.map((e: Record<string, unknown>) => e.UserId as number).filter(Boolean);
          } catch {
            // Fall back to provided userIds if PM_Employees doesn't exist
            targetUsers = userIds || [];
          }
          break;
      }

      return targetUsers;
    } catch (error) {
      logger.error('PolicyService', 'Failed to resolve target users:', error);
      return [];
    }
  }

  /**
   * Create distribution record
   */
  private async createDistribution(distribution: Partial<IPolicyDistribution>): Promise<IPolicyDistribution> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.POLICY_DISTRIBUTIONS_LIST)
        .items.add({
          Title: distribution.DistributionName,
          PolicyId: distribution.PolicyId,
          DistributionName: distribution.DistributionName,
          DistributionScope: distribution.DistributionScope,
          ScheduledDate: distribution.ScheduledDate?.toISOString(),
          DistributedDate: new Date().toISOString(),
          TargetCount: distribution.TargetCount,
          DueDate: distribution.DueDate?.toISOString(),
          TotalSent: distribution.TargetCount,
          TotalDelivered: 0,
          TotalOpened: 0,
          TotalAcknowledged: 0,
          TotalOverdue: 0,
          TotalExempted: 0,
          TotalFailed: 0,
          IsActive: true
        });

      return result.data as IPolicyDistribution;
    } catch (error) {
      logger.error('PolicyService', 'Failed to create distribution:', error);
      throw error;
    }
  }

  /**
   * Log audit entry
   */
  private async logAudit(audit: Partial<IPolicyAuditLog>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.POLICY_AUDIT_LOG_LIST)
        .items.add({
          Title: `${audit.Action} - ${audit.EntityType} ${audit.EntityId}`,
          EntityType: audit.EntityType,
          EntityId: audit.EntityId,
          PolicyId: audit.PolicyId,
          Action: audit.Action,
          ActionDescription: audit.ActionDescription,
          PerformedById: audit.PerformedById,
          PerformedByEmail: audit.PerformedByEmail,
          ActionDate: audit.ActionDate?.toISOString(),
          ComplianceRelevant: audit.ComplianceRelevant || false
        });
    } catch (error) {
      // Don't throw on audit log failures
      logger.error('PolicyService', 'Failed to log audit:', error);
    }
  }

  /**
   * Map SharePoint item to Policy interface
   */
  private mapPolicyItem(item: any): IPolicy {
    return {
      ...item,
      Tags: item.Tags ? JSON.parse(item.Tags) : [],
      RelatedPolicyIds: item.RelatedPolicyIds ? JSON.parse(item.RelatedPolicyIds) : [],
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      EffectiveDate: item.EffectiveDate ? new Date(item.EffectiveDate) : undefined,
      ExpiryDate: item.ExpiryDate ? new Date(item.ExpiryDate) : undefined,
      NextReviewDate: item.NextReviewDate ? new Date(item.NextReviewDate) : undefined
    } as IPolicy;
  }

  // ============================================================================
  // QUIZ METHODS (delegated from QuizService for convenience)
  // ============================================================================

  /**
   * Get quizzes associated with a policy
   */
  public async getQuizzesForPolicy(policyId: number): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(QuizLists.POLICY_QUIZZES)
        .items
        .filter(`PolicyId eq ${policyId}`)
        .select('Id', 'Title', 'PolicyId', 'PassingScore', 'TimeLimit', 'AllowRetake', 'MaxAttempts')();
      return items;
    } catch (error) {
      logger.error('PolicyService', 'Failed to get quizzes for policy', { policyId, error });
      return [];
    }
  }

  /**
   * Get questions for a quiz
   */
  public async getQuizQuestions(quizId: number): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(QuizLists.POLICY_QUIZ_QUESTIONS)
        .items
        .filter(`QuizId eq ${quizId}`)
        .select('Id', 'Title', 'QuizId', 'QuestionText', 'QuestionType', 'OptionA', 'OptionB', 'OptionC', 'OptionD', 'CorrectAnswer', 'Points', 'Explanation')
        .orderBy('SortOrder', true)();
      return items;
    } catch (error) {
      logger.error('PolicyService', 'Failed to get quiz questions', { quizId, error });
      return [];
    }
  }

  /**
   * Submit quiz result
   */
  public async submitQuizResult(result: {
    quizId: number;
    policyId: number;
    userId: number;
    score: number;
    passed: boolean;
    answers: string;
    attemptNumber: number;
  }): Promise<any> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(QuizLists.POLICY_QUIZ_RESULTS)
        .items
        .add({
          Title: `Quiz Result - ${new Date().toISOString()}`,
          QuizId: result.quizId,
          PolicyId: result.policyId,
          UserId: result.userId,
          Score: result.score,
          Passed: result.passed,
          Answers: result.answers,
          AttemptNumber: result.attemptNumber,
          CompletedDate: new Date().toISOString()
        });
      return item;
    } catch (error) {
      logger.error('PolicyService', 'Failed to submit quiz result', { result, error });
      throw error;
    }
  }
}
