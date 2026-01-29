// @ts-nocheck
/**
 * Policy Validation Service
 * Server-side validation for policy operations
 * Ensures actions are authorized regardless of UI restrictions
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import {
  IPolicy,
  IPolicyAcknowledgement,
  PolicyStatus,
  AcknowledgementStatus,
  ComplianceRisk,
  DataClassification
} from '../models/IPolicy';
import { logger } from './LoggingService';
import { PolicyLists } from '../constants/SharePointListNames';

// ============================================================================
// ENUMS & INTERFACES
// ============================================================================

/**
 * Policy operations that require validation
 */
export enum PolicyOperation {
  // Policy Lifecycle
  Create = 'Create',
  Read = 'Read',
  Update = 'Update',
  Delete = 'Delete',

  // Workflow Operations
  Submit = 'Submit',
  Approve = 'Approve',
  Reject = 'Reject',
  Publish = 'Publish',
  Archive = 'Archive',
  Retire = 'Retire',

  // Distribution
  Distribute = 'Distribute',
  Assign = 'Assign',

  // Acknowledgement
  Acknowledge = 'Acknowledge',
  Delegate = 'Delegate',
  Exempt = 'Exempt',

  // Administration
  BulkUpdate = 'BulkUpdate',
  Export = 'Export',
  ViewAudit = 'ViewAudit',
  ManageRetention = 'ManageRetention'
}

/**
 * User roles for policy management
 */
export enum PolicyRole {
  Employee = 'Employee',
  Author = 'Author',
  Reviewer = 'Reviewer',
  Approver = 'Approver',
  Publisher = 'Publisher',
  Administrator = 'Administrator',
  ComplianceOfficer = 'ComplianceOfficer',
  Auditor = 'Auditor'
}

/**
 * Validation result
 */
export interface IValidationResult {
  isValid: boolean;
  errors: IValidationError[];
  warnings: IValidationWarning[];
  context?: IValidationContext;
}

/**
 * Validation error
 */
export interface IValidationError {
  code: string;
  field?: string;
  message: string;
  severity: 'Critical' | 'Error';
}

/**
 * Validation warning
 */
export interface IValidationWarning {
  code: string;
  field?: string;
  message: string;
  suggestion?: string;
}

/**
 * Validation context for audit
 */
export interface IValidationContext {
  operation: PolicyOperation;
  userId: number;
  userEmail: string;
  userRoles: PolicyRole[];
  policyId?: number;
  timestamp: Date;
  ipAddress?: string;
  userAgent?: string;
}

/**
 * Permission matrix entry
 */
interface IPermissionMatrix {
  [role: string]: {
    [operation: string]: boolean | ((policy: IPolicy, userId: number) => boolean);
  };
}

// ============================================================================
// SERVICE CLASS
// ============================================================================

export class PolicyValidationService {
  private sp: SPFI;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserRoles: PolicyRole[] = [];
  private initialized: boolean = false;

  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly ACKNOWLEDGEMENTS_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;

  // Permission matrix defining what roles can do what operations
  private readonly permissionMatrix: IPermissionMatrix = {
    [PolicyRole.Employee]: {
      [PolicyOperation.Read]: true,
      [PolicyOperation.Acknowledge]: true,
      [PolicyOperation.Delegate]: false, // Can request delegation
    },
    [PolicyRole.Author]: {
      [PolicyOperation.Create]: true,
      [PolicyOperation.Read]: true,
      [PolicyOperation.Update]: (policy, userId) => policy.PolicyOwnerId === userId || policy.PolicyAuthorIds?.includes(userId),
      [PolicyOperation.Submit]: (policy, userId) => policy.PolicyOwnerId === userId || policy.PolicyAuthorIds?.includes(userId),
      [PolicyOperation.Acknowledge]: true,
    },
    [PolicyRole.Reviewer]: {
      [PolicyOperation.Read]: true,
      [PolicyOperation.Update]: (policy, userId) => policy.ReviewerIds?.includes(userId),
      [PolicyOperation.Acknowledge]: true,
    },
    [PolicyRole.Approver]: {
      [PolicyOperation.Read]: true,
      [PolicyOperation.Approve]: (policy, userId) => policy.ApproverIds?.includes(userId),
      [PolicyOperation.Reject]: (policy, userId) => policy.ApproverIds?.includes(userId),
      [PolicyOperation.Acknowledge]: true,
    },
    [PolicyRole.Publisher]: {
      [PolicyOperation.Read]: true,
      [PolicyOperation.Publish]: true,
      [PolicyOperation.Distribute]: true,
      [PolicyOperation.Archive]: true,
      [PolicyOperation.Acknowledge]: true,
    },
    [PolicyRole.Administrator]: {
      [PolicyOperation.Create]: true,
      [PolicyOperation.Read]: true,
      [PolicyOperation.Update]: true,
      [PolicyOperation.Delete]: true,
      [PolicyOperation.Submit]: true,
      [PolicyOperation.Approve]: true,
      [PolicyOperation.Reject]: true,
      [PolicyOperation.Publish]: true,
      [PolicyOperation.Archive]: true,
      [PolicyOperation.Retire]: true,
      [PolicyOperation.Distribute]: true,
      [PolicyOperation.Assign]: true,
      [PolicyOperation.Acknowledge]: true,
      [PolicyOperation.Delegate]: true,
      [PolicyOperation.Exempt]: true,
      [PolicyOperation.BulkUpdate]: true,
      [PolicyOperation.Export]: true,
      [PolicyOperation.ViewAudit]: true,
      [PolicyOperation.ManageRetention]: true,
    },
    [PolicyRole.ComplianceOfficer]: {
      [PolicyOperation.Read]: true,
      [PolicyOperation.Approve]: true,
      [PolicyOperation.Exempt]: true,
      [PolicyOperation.ViewAudit]: true,
      [PolicyOperation.Export]: true,
      [PolicyOperation.ManageRetention]: true,
    },
    [PolicyRole.Auditor]: {
      [PolicyOperation.Read]: true,
      [PolicyOperation.ViewAudit]: true,
      [PolicyOperation.Export]: true,
    },
  };

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // INITIALIZATION
  // ============================================================================

  /**
   * Initialize the service and determine user roles
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;

      // Determine user roles based on group membership
      this.currentUserRoles = await this.getUserRoles(user.Id);
      this.initialized = true;

      logger.info('PolicyValidationService', `Initialized for user ${user.Email} with roles: ${this.currentUserRoles.join(', ')}`);
    } catch (error) {
      logger.error('PolicyValidationService', 'Failed to initialize:', error);
      throw error;
    }
  }

  /**
   * Get user roles based on SharePoint group membership
   */
  private async getUserRoles(userId: number): Promise<PolicyRole[]> {
    const roles: PolicyRole[] = [PolicyRole.Employee]; // Everyone starts as employee

    try {
      const groups = await this.sp.web.currentUser.groups();

      const roleGroupMapping: { [groupPattern: string]: PolicyRole } = {
        'policy author': PolicyRole.Author,
        'policy reviewer': PolicyRole.Reviewer,
        'policy approver': PolicyRole.Approver,
        'policy publisher': PolicyRole.Publisher,
        'policy admin': PolicyRole.Administrator,
        'compliance officer': PolicyRole.ComplianceOfficer,
        'auditor': PolicyRole.Auditor,
        'site admin': PolicyRole.Administrator,
        'owner': PolicyRole.Administrator,
      };

      for (const group of groups) {
        const groupNameLower = group.Title.toLowerCase();
        for (const [pattern, role] of Object.entries(roleGroupMapping)) {
          if (groupNameLower.includes(pattern) && !roles.includes(role)) {
            roles.push(role);
          }
        }
      }
    } catch (error) {
      logger.warn('PolicyValidationService', 'Failed to get user groups:', error);
    }

    return roles;
  }

  // ============================================================================
  // PERMISSION VALIDATION
  // ============================================================================

  /**
   * Check if current user can perform an operation on a policy
   */
  public async canPerformOperation(
    operation: PolicyOperation,
    policyId?: number,
    additionalContext?: Partial<IValidationContext>
  ): Promise<IValidationResult> {
    await this.initialize();

    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    // Get policy if ID provided
    let policy: IPolicy | null = null;
    if (policyId) {
      try {
        policy = await this.sp.web.lists
          .getByTitle(this.POLICIES_LIST)
          .items.getById(policyId)() as IPolicy;
      } catch {
        errors.push({
          code: 'POLICY_NOT_FOUND',
          message: `Policy with ID ${policyId} not found`,
          severity: 'Critical'
        });
        return { isValid: false, errors, warnings };
      }
    }

    // Check if any role grants permission
    let hasPermission = false;
    for (const role of this.currentUserRoles) {
      const rolePerms = this.permissionMatrix[role];
      if (!rolePerms) continue;

      const permission = rolePerms[operation];
      if (permission === true) {
        hasPermission = true;
        break;
      } else if (typeof permission === 'function' && policy) {
        if (permission(policy, this.currentUserId)) {
          hasPermission = true;
          break;
        }
      }
    }

    if (!hasPermission) {
      errors.push({
        code: 'INSUFFICIENT_PERMISSIONS',
        message: `User does not have permission to perform ${operation}`,
        severity: 'Critical'
      });
    }

    // Additional validation based on operation
    if (hasPermission && policy) {
      const operationValidation = await this.validateOperationContext(operation, policy);
      errors.push(...operationValidation.errors);
      warnings.push(...operationValidation.warnings);
    }

    const context: IValidationContext = {
      operation,
      userId: this.currentUserId,
      userEmail: this.currentUserEmail,
      userRoles: this.currentUserRoles,
      policyId,
      timestamp: new Date(),
      ...additionalContext
    };

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      context
    };
  }

  /**
   * Validate operation context (status transitions, etc.)
   */
  private async validateOperationContext(
    operation: PolicyOperation,
    policy: IPolicy
  ): Promise<{ errors: IValidationError[]; warnings: IValidationWarning[] }> {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    switch (operation) {
      case PolicyOperation.Submit:
        if (policy.PolicyStatus !== PolicyStatus.Draft) {
          errors.push({
            code: 'INVALID_STATUS_TRANSITION',
            message: 'Only draft policies can be submitted for review',
            severity: 'Error'
          });
        }
        break;

      case PolicyOperation.Approve:
        if (policy.PolicyStatus !== PolicyStatus.InReview && policy.PolicyStatus !== PolicyStatus.PendingApproval) {
          errors.push({
            code: 'INVALID_STATUS_TRANSITION',
            message: 'Only policies in review or pending approval can be approved',
            severity: 'Error'
          });
        }
        break;

      case PolicyOperation.Reject:
        if (policy.PolicyStatus !== PolicyStatus.InReview && policy.PolicyStatus !== PolicyStatus.PendingApproval) {
          errors.push({
            code: 'INVALID_STATUS_TRANSITION',
            message: 'Only policies in review or pending approval can be rejected',
            severity: 'Error'
          });
        }
        break;

      case PolicyOperation.Publish:
        if (policy.PolicyStatus !== PolicyStatus.Approved) {
          errors.push({
            code: 'INVALID_STATUS_TRANSITION',
            message: 'Only approved policies can be published',
            severity: 'Error'
          });
        }
        // Check for required fields
        if (!policy.EffectiveDate) {
          warnings.push({
            code: 'MISSING_EFFECTIVE_DATE',
            message: 'Policy has no effective date set',
            suggestion: 'Set an effective date before publishing'
          });
        }
        break;

      case PolicyOperation.Archive:
        if (policy.PolicyStatus !== PolicyStatus.Published && policy.PolicyStatus !== PolicyStatus.Expired) {
          errors.push({
            code: 'INVALID_STATUS_TRANSITION',
            message: 'Only published or expired policies can be archived',
            severity: 'Error'
          });
        }
        // Check for pending acknowledgements
        const pendingAcks = await this.sp.web.lists
          .getByTitle(this.ACKNOWLEDGEMENTS_LIST)
          .items.filter(`PolicyId eq ${policy.Id} and AckStatus ne '${AcknowledgementStatus.Acknowledged}'`)
          .top(1)();
        if (pendingAcks.length > 0) {
          warnings.push({
            code: 'PENDING_ACKNOWLEDGEMENTS',
            message: 'Policy has pending acknowledgements',
            suggestion: 'Consider notifying affected users before archiving'
          });
        }
        break;

      case PolicyOperation.Delete:
        if (policy.PolicyStatus === PolicyStatus.Published) {
          errors.push({
            code: 'CANNOT_DELETE_PUBLISHED',
            message: 'Published policies cannot be deleted. Archive first.',
            severity: 'Error'
          });
        }
        break;

      case PolicyOperation.Update:
        if (policy.PolicyStatus === PolicyStatus.Archived || policy.PolicyStatus === PolicyStatus.Retired) {
          errors.push({
            code: 'CANNOT_UPDATE_ARCHIVED',
            message: 'Archived or retired policies cannot be updated',
            severity: 'Error'
          });
        }
        break;
    }

    return { errors, warnings };
  }

  // ============================================================================
  // DATA VALIDATION
  // ============================================================================

  /**
   * Validate policy data before save
   */
  public validatePolicyData(policy: Partial<IPolicy>): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    // Required fields
    if (!policy.PolicyName?.trim()) {
      errors.push({
        code: 'REQUIRED_FIELD',
        field: 'PolicyName',
        message: 'Policy name is required',
        severity: 'Error'
      });
    } else if (policy.PolicyName.length > 255) {
      errors.push({
        code: 'FIELD_TOO_LONG',
        field: 'PolicyName',
        message: 'Policy name cannot exceed 255 characters',
        severity: 'Error'
      });
    }

    if (!policy.PolicyCategory) {
      errors.push({
        code: 'REQUIRED_FIELD',
        field: 'PolicyCategory',
        message: 'Policy category is required',
        severity: 'Error'
      });
    }

    if (!policy.PolicyOwnerId) {
      errors.push({
        code: 'REQUIRED_FIELD',
        field: 'PolicyOwnerId',
        message: 'Policy owner is required',
        severity: 'Error'
      });
    }

    // Business rules
    if (policy.RequiresQuiz && !policy.QuizPassingScore) {
      warnings.push({
        code: 'QUIZ_NO_PASSING_SCORE',
        field: 'QuizPassingScore',
        message: 'Quiz is required but no passing score is set',
        suggestion: 'Set a passing score (e.g., 80%)'
      });
    }

    if (policy.RequiresAcknowledgement && !policy.AcknowledgementDeadlineDays && !policy.ReadTimeframe) {
      warnings.push({
        code: 'NO_ACKNOWLEDGEMENT_DEADLINE',
        field: 'AcknowledgementDeadlineDays',
        message: 'Acknowledgement required but no deadline set',
        suggestion: 'Set a deadline for acknowledgements'
      });
    }

    // Date validations
    if (policy.EffectiveDate && policy.ExpiryDate) {
      const effective = new Date(policy.EffectiveDate);
      const expiry = new Date(policy.ExpiryDate);
      if (expiry <= effective) {
        errors.push({
          code: 'INVALID_DATE_RANGE',
          field: 'ExpiryDate',
          message: 'Expiry date must be after effective date',
          severity: 'Error'
        });
      }
    }

    // Compliance risk validation
    if (policy.ComplianceRisk === ComplianceRisk.Critical || policy.ComplianceRisk === ComplianceRisk.High) {
      if (!policy.RequiresAcknowledgement) {
        warnings.push({
          code: 'HIGH_RISK_NO_ACKNOWLEDGEMENT',
          field: 'RequiresAcknowledgement',
          message: 'High/Critical risk policy should require acknowledgement',
          suggestion: 'Enable acknowledgement requirement'
        });
      }
    }

    // Data classification validation
    if (policy.DataClassification === DataClassification.Restricted ||
        policy.DataClassification === DataClassification.Confidential) {
      if (!policy.TargetDepartments?.length && !policy.TargetRoles?.length &&
          !policy.TargetUserIds?.length) {
        warnings.push({
          code: 'SENSITIVE_NO_TARGETING',
          field: 'DataClassification',
          message: 'Sensitive policy has no distribution targeting',
          suggestion: 'Consider restricting distribution to specific audiences'
        });
      }
    }

    return {
      isValid: errors.filter(e => e.severity === 'Error' || e.severity === 'Critical').length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate acknowledgement request
   */
  public async validateAcknowledgement(
    acknowledgementId: number,
    userId?: number
  ): Promise<IValidationResult> {
    await this.initialize();

    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];
    const targetUserId = userId || this.currentUserId;

    try {
      // Get acknowledgement
      const ack = await this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENTS_LIST)
        .items.getById(acknowledgementId)() as IPolicyAcknowledgement;

      // Verify user owns this acknowledgement
      if (ack.AckUserId !== targetUserId) {
        // Check if delegated
        if (!ack.IsDelegated || ack.DelegatedById !== targetUserId) {
          errors.push({
            code: 'UNAUTHORIZED_ACKNOWLEDGEMENT',
            message: 'User is not authorized to acknowledge this policy',
            severity: 'Critical'
          });
        }
      }

      // Check if already acknowledged
      if (ack.AckStatus === AcknowledgementStatus.Acknowledged) {
        errors.push({
          code: 'ALREADY_ACKNOWLEDGED',
          message: 'Policy has already been acknowledged',
          severity: 'Error'
        });
      }

      // Check if exempted
      if (ack.IsExempted) {
        warnings.push({
          code: 'EXEMPTED_POLICY',
          message: 'User is exempted from this policy',
          suggestion: 'Acknowledgement will clear the exemption'
        });
      }

      // Check quiz requirement
      if (ack.QuizRequired && ack.QuizStatus !== 'Passed') {
        errors.push({
          code: 'QUIZ_NOT_PASSED',
          message: 'Quiz must be passed before acknowledgement',
          severity: 'Error'
        });
      }

      // Check if policy still active
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(ack.PolicyId)() as IPolicy;

      if (!policy.IsActive || policy.PolicyStatus === PolicyStatus.Archived ||
          policy.PolicyStatus === PolicyStatus.Retired) {
        errors.push({
          code: 'POLICY_INACTIVE',
          message: 'Policy is no longer active',
          severity: 'Error'
        });
      }

    } catch (error) {
      errors.push({
        code: 'ACKNOWLEDGEMENT_NOT_FOUND',
        message: `Acknowledgement with ID ${acknowledgementId} not found`,
        severity: 'Critical'
      });
    }

    return { isValid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate delegation request
   */
  public async validateDelegation(
    policyId: number,
    delegateToUserId: number,
    reason: string
  ): Promise<IValidationResult> {
    await this.initialize();

    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    // Check if policy allows delegation
    try {
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)() as IPolicy;

      // Critical/High risk policies may not allow delegation
      if (policy.ComplianceRisk === ComplianceRisk.Critical) {
        errors.push({
          code: 'DELEGATION_NOT_ALLOWED',
          message: 'Critical risk policies cannot be delegated',
          severity: 'Critical'
        });
      }

      // Check if delegate user exists
      try {
        await this.sp.web.siteUsers.getById(delegateToUserId)();
      } catch {
        errors.push({
          code: 'INVALID_DELEGATE',
          message: 'Delegate user not found',
          severity: 'Error'
        });
      }

      // Reason is required
      if (!reason?.trim()) {
        errors.push({
          code: 'REQUIRED_FIELD',
          field: 'reason',
          message: 'Delegation reason is required',
          severity: 'Error'
        });
      }

    } catch {
      errors.push({
        code: 'POLICY_NOT_FOUND',
        message: `Policy with ID ${policyId} not found`,
        severity: 'Critical'
      });
    }

    return { isValid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate bulk operation
   */
  public async validateBulkOperation(
    operation: PolicyOperation,
    policyIds: number[]
  ): Promise<IValidationResult> {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    // Check bulk operation permission
    const canBulk = await this.canPerformOperation(PolicyOperation.BulkUpdate);
    if (!canBulk.isValid) {
      return canBulk;
    }

    // Validate each policy
    const failedPolicies: number[] = [];
    for (const policyId of policyIds) {
      const result = await this.canPerformOperation(operation, policyId);
      if (!result.isValid) {
        failedPolicies.push(policyId);
      }
    }

    if (failedPolicies.length > 0) {
      errors.push({
        code: 'PARTIAL_BULK_FAILURE',
        message: `Operation cannot be performed on ${failedPolicies.length} of ${policyIds.length} policies`,
        severity: 'Error'
      });
    }

    if (policyIds.length > 100) {
      warnings.push({
        code: 'LARGE_BULK_OPERATION',
        message: 'Large bulk operation may take significant time',
        suggestion: 'Consider processing in smaller batches'
      });
    }

    return { isValid: errors.length === 0, errors, warnings };
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get current user's roles
   */
  public getCurrentUserRoles(): PolicyRole[] {
    return [...this.currentUserRoles];
  }

  /**
   * Get current user ID
   */
  public getCurrentUserId(): number {
    return this.currentUserId;
  }

  /**
   * Check if user has specific role
   */
  public hasRole(role: PolicyRole): boolean {
    return this.currentUserRoles.includes(role);
  }

  /**
   * Check if user is administrator
   */
  public isAdministrator(): boolean {
    return this.currentUserRoles.includes(PolicyRole.Administrator);
  }

  /**
   * Check if user is compliance officer
   */
  public isComplianceOfficer(): boolean {
    return this.currentUserRoles.includes(PolicyRole.ComplianceOfficer);
  }

  /**
   * Format validation result for display
   */
  public formatValidationResult(result: IValidationResult): string {
    const lines: string[] = [];

    if (result.isValid) {
      lines.push('Validation passed.');
    } else {
      lines.push('Validation failed:');
    }

    for (const error of result.errors) {
      lines.push(`  [${error.severity}] ${error.code}: ${error.message}`);
    }

    for (const warning of result.warnings) {
      lines.push(`  [Warning] ${warning.code}: ${warning.message}`);
      if (warning.suggestion) {
        lines.push(`    Suggestion: ${warning.suggestion}`);
      }
    }

    return lines.join('\n');
  }
}
