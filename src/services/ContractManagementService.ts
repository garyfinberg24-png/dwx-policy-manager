// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix type mismatches with enum types and interface properties
/**
 * Contract Management Service
 *
 * Comprehensive service for enterprise contract lifecycle management.
 * Handles contracts, clauses, approvals, signatures, obligations, and audit logging.
 * Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import '@pnp/sp/files';
import '@pnp/sp/folders';

import {
  IContractRecord,
  IContractParty,
  IContractVersion,
  IContractClause,
  IContractClauseInstance,
  IContractTemplate,
  IContractApproval,
  IContractApprovalRule,
  IContractSignature,
  IContractObligation,
  IContractAuditLog,
  IContractComment,
  IContractDocument,
  IContractStatistics,
  IContractDashboard,
  IContractFilter,
  IClauseFilter,
  IObligationFilter,
  IContractAlert,
  IExpiryTimelineItem,
  ContractLifecycleStatus,
  ContractCategory,
  ContractPriority,
  ContractRiskLevel,
  ContractValueType,
  ContractRenewalType,
  ContractApprovalStatus,
  ContractApprovalAction,
  SignatureStatus,
  ObligationStatus,
  ObligationFrequency,
  ContractAuditAction,
  ClauseCategory,
  ClauseRiskLevel
} from '../models/IContractManagement';
import { Currency, PaymentTerms } from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

// ==================== LIST NAMES ====================

const LISTS = {
  CONTRACTS: 'JML_ContractRecords',
  PARTIES: 'JML_ContractParties',
  VERSIONS: 'JML_ContractVersions',
  CLAUSES: 'JML_ContractClauses',
  CLAUSE_INSTANCES: 'JML_ContractClauseInstances',
  CLAUSE_CATEGORIES: 'JML_ClauseCategories',
  TEMPLATES: 'JML_ContractTemplates',
  APPROVALS: 'JML_ContractApprovals',
  APPROVAL_RULES: 'JML_ContractApprovalRules',
  SIGNATURES: 'JML_ContractSignatures',
  OBLIGATIONS: 'JML_ContractObligations',
  AUDIT_LOG: 'JML_ContractAuditLog',
  COMMENTS: 'JML_ContractComments',
  DOCUMENTS: 'JML_ContractDocuments'
};

// ==================== MAIN SERVICE CLASS ====================

export class ContractManagementService {
  private sp: SPFI;
  private currentUserId: number | null = null;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== INITIALIZATION ====================

  public async initialize(): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();
      this.currentUserId = currentUser.Id;
    } catch (error) {
      logger.error('ContractManagementService', 'Error initializing service:', error);
    }
  }

  // ==================== CONTRACT CRUD ====================

  /**
   * Get contracts with optional filtering
   */
  public async getContracts(filter?: IContractFilter): Promise<IContractRecord[]> {
    try {
      let query = this.sp.web.lists.getByTitle(LISTS.CONTRACTS).items
        .select(
          'Id', 'Title', 'ContractNumber', 'ExternalReference', 'Category', 'Status',
          'Priority', 'RiskLevel', 'Industry', 'Description', 'ExecutiveSummary',
          'EffectiveDate', 'ExpirationDate', 'SignedDate', 'TerminationDate',
          'RenewalType', 'RenewalTermMonths', 'RenewalNotificationDays', 'NextRenewalDate',
          'TerminationNoticeDays', 'ValueType', 'TotalValue', 'AnnualValue', 'MonthlyValue',
          'Currency', 'PaymentTerms', 'BudgetCode', 'CostCenter',
          'OwnerId', 'Owner/Title', 'Owner/EMail', 'SecondaryOwnerIds', 'Department', 'BusinessUnit',
          'PrimaryCounterpartyId', 'CounterpartyName', 'CounterpartyContact', 'CounterpartyEmail',
          'DocumentLibraryUrl', 'CurrentVersionUrl', 'ExecutedDocumentUrl',
          'Version', 'IsAmendment', 'ParentContractId', 'AmendmentReason',
          'TemplateId', 'TemplateName', 'ComplianceRequirements', 'DataClassification',
          'GDPRApplicable', 'RequiresLegalReview', 'RiskScore', 'RiskFactors',
          'HasSLATerms', 'Tags', 'Notes', 'CurrentApprovalStage', 'TotalApprovalStages',
          'SubmittedForApprovalDate', 'ApprovedDate', 'SentForSignatureDate', 'FullyExecutedDate',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Owner', 'Author', 'Editor');

      // Apply filters
      const filters = this.buildContractFilters(filter);
      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.orderBy('ExpirationDate', true).top(5000)();
      return items.map(this.mapContractFromSP);
    } catch (error) {
      logger.error('ContractManagementService', 'Error getting contracts:', error);
      throw error;
    }
  }

  /**
   * Get a single contract by ID
   */
  public async getContractById(id: number): Promise<IContractRecord> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(LISTS.CONTRACTS).items
        .getById(validId)
        .select(
          'Id', 'Title', 'ContractNumber', 'ExternalReference', 'Category', 'Status',
          'Priority', 'RiskLevel', 'Industry', 'Description', 'ExecutiveSummary',
          'EffectiveDate', 'ExpirationDate', 'SignedDate', 'TerminationDate',
          'RenewalType', 'RenewalTermMonths', 'RenewalNotificationDays', 'NextRenewalDate',
          'TerminationNoticeDays', 'ValueType', 'TotalValue', 'AnnualValue', 'MonthlyValue',
          'Currency', 'PaymentTerms', 'BudgetCode', 'CostCenter',
          'OwnerId', 'Owner/Title', 'Owner/EMail', 'SecondaryOwnerIds', 'Department', 'BusinessUnit',
          'PrimaryCounterpartyId', 'CounterpartyName', 'CounterpartyContact', 'CounterpartyEmail',
          'DocumentLibraryUrl', 'CurrentVersionUrl', 'ExecutedDocumentUrl',
          'Version', 'IsAmendment', 'ParentContractId', 'AmendmentReason',
          'TemplateId', 'TemplateName', 'ComplianceRequirements', 'DataClassification',
          'GDPRApplicable', 'RequiresLegalReview', 'RiskScore', 'RiskFactors',
          'HasSLATerms', 'Tags', 'Notes', 'CurrentApprovalStage', 'TotalApprovalStages',
          'SubmittedForApprovalDate', 'ApprovedDate', 'SentForSignatureDate', 'FullyExecutedDate',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Owner', 'Author', 'Editor')();

      return this.mapContractFromSP(item);
    } catch (error) {
      logger.error('ContractManagementService', `Error getting contract ${id}:`, error);
      throw error;
    }
  }

  /**
   * Create a new contract
   */
  public async createContract(contract: Partial<IContractRecord>): Promise<IContractRecord> {
    try {
      // Validate required fields
      if (!contract.Title) {
        throw new Error('Contract title is required');
      }

      // Generate contract number
      const contractNumber = contract.ContractNumber || await this.generateContractNumber();

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(contract.Title),
        ContractNumber: contractNumber,
        Category: contract.Category || ContractCategory.Other,
        Status: contract.Status || ContractLifecycleStatus.Draft,
        Priority: contract.Priority || ContractPriority.Medium,
        RiskLevel: contract.RiskLevel || ContractRiskLevel.Medium,
        Industry: contract.Industry,
        Description: contract.Description ? ValidationUtils.sanitizeHtml(contract.Description) : null,
        ExecutiveSummary: contract.ExecutiveSummary ? ValidationUtils.sanitizeHtml(contract.ExecutiveSummary) : null,
        EffectiveDate: contract.EffectiveDate,
        ExpirationDate: contract.ExpirationDate,
        RenewalType: contract.RenewalType || ContractRenewalType.ManualRenew,
        RenewalTermMonths: contract.RenewalTermMonths || 12,
        RenewalNotificationDays: contract.RenewalNotificationDays || 90,
        TerminationNoticeDays: contract.TerminationNoticeDays || 30,
        ValueType: contract.ValueType || ContractValueType.FixedFee,
        TotalValue: contract.TotalValue || 0,
        AnnualValue: contract.AnnualValue,
        MonthlyValue: contract.MonthlyValue,
        Currency: contract.Currency || Currency.ZAR,
        PaymentTerms: contract.PaymentTerms || PaymentTerms.Net30,
        BudgetCode: contract.BudgetCode,
        CostCenter: contract.CostCenter,
        OwnerId: contract.OwnerId || this.currentUserId,
        SecondaryOwnerIds: contract.SecondaryOwnerIds,
        Department: contract.Department,
        BusinessUnit: contract.BusinessUnit,
        PrimaryCounterpartyId: contract.PrimaryCounterpartyId,
        CounterpartyName: contract.CounterpartyName,
        CounterpartyContact: contract.CounterpartyContact,
        CounterpartyEmail: contract.CounterpartyEmail,
        Version: 1,
        IsAmendment: contract.IsAmendment || false,
        ParentContractId: contract.ParentContractId,
        AmendmentReason: contract.AmendmentReason,
        TemplateId: contract.TemplateId,
        TemplateName: contract.TemplateName,
        ComplianceRequirements: contract.ComplianceRequirements,
        DataClassification: contract.DataClassification,
        GDPRApplicable: contract.GDPRApplicable || false,
        RequiresLegalReview: contract.RequiresLegalReview || false,
        RiskScore: contract.RiskScore,
        RiskFactors: contract.RiskFactors,
        HasSLATerms: contract.HasSLATerms || false,
        Tags: contract.Tags,
        Notes: contract.Notes
      };

      const result = await this.sp.web.lists.getByTitle(LISTS.CONTRACTS).items.add(itemData);

      // Log audit
      await this.logAudit(result.data.Id, ContractAuditAction.Created, 'Contract created', 'Contract');

      return this.getContractById(result.data.Id);
    } catch (error) {
      logger.error('ContractManagementService', 'Error creating contract:', error);
      throw error;
    }
  }

  /**
   * Update an existing contract
   */
  public async updateContract(id: number, updates: Partial<IContractRecord>): Promise<IContractRecord> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Get current state for audit
      const currentContract = await this.getContractById(validId);

      const itemData: Record<string, unknown> = {};

      // Map updates to SharePoint fields
      if (updates.Title !== undefined) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Category !== undefined) itemData.Category = updates.Category;
      if (updates.Status !== undefined) itemData.Status = updates.Status;
      if (updates.Priority !== undefined) itemData.Priority = updates.Priority;
      if (updates.RiskLevel !== undefined) itemData.RiskLevel = updates.RiskLevel;
      if (updates.Industry !== undefined) itemData.Industry = updates.Industry;
      if (updates.Description !== undefined) itemData.Description = updates.Description ? ValidationUtils.sanitizeHtml(updates.Description) : null;
      if (updates.ExecutiveSummary !== undefined) itemData.ExecutiveSummary = updates.ExecutiveSummary;
      if (updates.EffectiveDate !== undefined) itemData.EffectiveDate = updates.EffectiveDate;
      if (updates.ExpirationDate !== undefined) itemData.ExpirationDate = updates.ExpirationDate;
      if (updates.SignedDate !== undefined) itemData.SignedDate = updates.SignedDate;
      if (updates.TerminationDate !== undefined) itemData.TerminationDate = updates.TerminationDate;
      if (updates.RenewalType !== undefined) itemData.RenewalType = updates.RenewalType;
      if (updates.RenewalTermMonths !== undefined) itemData.RenewalTermMonths = updates.RenewalTermMonths;
      if (updates.RenewalNotificationDays !== undefined) itemData.RenewalNotificationDays = updates.RenewalNotificationDays;
      if (updates.NextRenewalDate !== undefined) itemData.NextRenewalDate = updates.NextRenewalDate;
      if (updates.TerminationNoticeDays !== undefined) itemData.TerminationNoticeDays = updates.TerminationNoticeDays;
      if (updates.ValueType !== undefined) itemData.ValueType = updates.ValueType;
      if (updates.TotalValue !== undefined) itemData.TotalValue = updates.TotalValue;
      if (updates.AnnualValue !== undefined) itemData.AnnualValue = updates.AnnualValue;
      if (updates.MonthlyValue !== undefined) itemData.MonthlyValue = updates.MonthlyValue;
      if (updates.Currency !== undefined) itemData.Currency = updates.Currency;
      if (updates.PaymentTerms !== undefined) itemData.PaymentTerms = updates.PaymentTerms;
      if (updates.BudgetCode !== undefined) itemData.BudgetCode = updates.BudgetCode;
      if (updates.CostCenter !== undefined) itemData.CostCenter = updates.CostCenter;
      if (updates.OwnerId !== undefined) itemData.OwnerId = updates.OwnerId;
      if (updates.SecondaryOwnerIds !== undefined) itemData.SecondaryOwnerIds = updates.SecondaryOwnerIds;
      if (updates.Department !== undefined) itemData.Department = updates.Department;
      if (updates.BusinessUnit !== undefined) itemData.BusinessUnit = updates.BusinessUnit;
      if (updates.PrimaryCounterpartyId !== undefined) itemData.PrimaryCounterpartyId = updates.PrimaryCounterpartyId;
      if (updates.CounterpartyName !== undefined) itemData.CounterpartyName = updates.CounterpartyName;
      if (updates.CounterpartyContact !== undefined) itemData.CounterpartyContact = updates.CounterpartyContact;
      if (updates.CounterpartyEmail !== undefined) itemData.CounterpartyEmail = updates.CounterpartyEmail;
      if (updates.DocumentLibraryUrl !== undefined) itemData.DocumentLibraryUrl = updates.DocumentLibraryUrl;
      if (updates.CurrentVersionUrl !== undefined) itemData.CurrentVersionUrl = updates.CurrentVersionUrl;
      if (updates.ExecutedDocumentUrl !== undefined) itemData.ExecutedDocumentUrl = updates.ExecutedDocumentUrl;
      if (updates.ComplianceRequirements !== undefined) itemData.ComplianceRequirements = updates.ComplianceRequirements;
      if (updates.DataClassification !== undefined) itemData.DataClassification = updates.DataClassification;
      if (updates.GDPRApplicable !== undefined) itemData.GDPRApplicable = updates.GDPRApplicable;
      if (updates.RequiresLegalReview !== undefined) itemData.RequiresLegalReview = updates.RequiresLegalReview;
      if (updates.RiskScore !== undefined) itemData.RiskScore = updates.RiskScore;
      if (updates.RiskFactors !== undefined) itemData.RiskFactors = updates.RiskFactors;
      if (updates.HasSLATerms !== undefined) itemData.HasSLATerms = updates.HasSLATerms;
      if (updates.Tags !== undefined) itemData.Tags = updates.Tags;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes;

      await this.sp.web.lists.getByTitle(LISTS.CONTRACTS).items.getById(validId).update(itemData);

      // Determine audit action
      let auditAction = ContractAuditAction.Updated;
      let auditDescription = 'Contract updated';

      if (updates.Status && updates.Status !== currentContract.Status) {
        auditAction = ContractAuditAction.StatusChanged;
        auditDescription = `Status changed from ${currentContract.Status} to ${updates.Status}`;
      }

      // Log audit
      await this.logAudit(
        validId,
        auditAction,
        auditDescription,
        'Contract',
        JSON.stringify(currentContract),
        JSON.stringify(updates)
      );

      return this.getContractById(validId);
    } catch (error) {
      logger.error('ContractManagementService', `Error updating contract ${id}:`, error);
      throw error;
    }
  }

  /**
   * Delete a contract (only drafts can be deleted)
   */
  public async deleteContract(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const contract = await this.getContractById(validId);
      if (contract.Status !== ContractLifecycleStatus.Draft) {
        throw new Error('Only draft contracts can be deleted');
      }

      // Delete related records first
      await this.deleteContractRelatedRecords(validId);

      // Delete the contract
      await this.sp.web.lists.getByTitle(LISTS.CONTRACTS).items.getById(validId).delete();
    } catch (error) {
      logger.error('ContractManagementService', `Error deleting contract ${id}:`, error);
      throw error;
    }
  }

  // ==================== CONTRACT WORKFLOW ====================

  /**
   * Submit contract for approval
   */
  public async submitForApproval(contractId: number, comments?: string): Promise<void> {
    try {
      const contract = await this.getContractById(contractId);

      if (contract.Status !== ContractLifecycleStatus.Draft && contract.Status !== ContractLifecycleStatus.InReview) {
        throw new Error('Contract must be in Draft or In Review status to submit for approval');
      }

      // Get applicable approval rules
      const rules = await this.getApplicableApprovalRules(contract);

      if (rules.length === 0) {
        throw new Error('No approval rules configured for this contract type');
      }

      // Create approval requests
      for (const rule of rules) {
        const approverIds = JSON.parse(rule.ApproverIds || '[]') as number[];
        for (const approverId of approverIds) {
          await this.createApprovalRequest(contractId, approverId, rule.ApprovalOrder, comments);
        }
      }

      // Update contract status
      await this.updateContract(contractId, {
        Status: ContractLifecycleStatus.PendingApproval,
        SubmittedForApprovalDate: new Date(),
        CurrentApprovalStage: 1,
        TotalApprovalStages: rules.length
      });

      // Log audit
      await this.logAudit(contractId, ContractAuditAction.ApprovalSubmitted, 'Contract submitted for approval', 'Approval');
    } catch (error) {
      logger.error('ContractManagementService', `Error submitting contract ${contractId} for approval:`, error);
      throw error;
    }
  }

  /**
   * Approve a contract
   */
  public async approveContract(approvalId: number, comments?: string): Promise<void> {
    try {
      const approval = await this.getApprovalById(approvalId);

      if (approval.Status !== ContractApprovalStatus.Pending) {
        throw new Error('This approval request is no longer pending');
      }

      // Update approval
      await this.sp.web.lists.getByTitle(LISTS.APPROVALS).items.getById(approvalId).update({
        Status: ContractApprovalStatus.Approved,
        Action: ContractApprovalAction.Approve,
        ActionDate: new Date(),
        ApprovalComments: comments
      });

      // Check if all approvals are complete
      const allApprovals = await this.getContractApprovals(approval.ContractId);
      const pendingApprovals = allApprovals.filter(a => a.Status === ContractApprovalStatus.Pending);

      if (pendingApprovals.length === 0) {
        // All approved - move to next stage or complete
        const contract = await this.getContractById(approval.ContractId);
        const currentStage = contract.CurrentApprovalStage || 1;
        const totalStages = contract.TotalApprovalStages || 1;

        if (currentStage >= totalStages) {
          // All stages complete
          await this.updateContract(approval.ContractId, {
            Status: ContractLifecycleStatus.Approved,
            ApprovedDate: new Date()
          });
        } else {
          // Move to next stage
          await this.updateContract(approval.ContractId, {
            CurrentApprovalStage: currentStage + 1
          });
          // TODO: Create next stage approvals
        }
      }

      // Log audit
      await this.logAudit(approval.ContractId, ContractAuditAction.ApprovalGranted, `Approved by user`, 'Approval');
    } catch (error) {
      logger.error('ContractManagementService', `Error approving contract:`, error);
      throw error;
    }
  }

  /**
   * Reject a contract
   */
  public async rejectContract(approvalId: number, reason: string): Promise<void> {
    try {
      if (!reason || reason.trim().length === 0) {
        throw new Error('Rejection reason is required');
      }

      const approval = await this.getApprovalById(approvalId);

      if (approval.Status !== ContractApprovalStatus.Pending) {
        throw new Error('This approval request is no longer pending');
      }

      // Update approval
      await this.sp.web.lists.getByTitle(LISTS.APPROVALS).items.getById(approvalId).update({
        Status: ContractApprovalStatus.Rejected,
        Action: ContractApprovalAction.Reject,
        ActionDate: new Date(),
        ApprovalComments: reason
      });

      // Update contract status back to draft
      await this.updateContract(approval.ContractId, {
        Status: ContractLifecycleStatus.Draft
      });

      // Cancel other pending approvals
      const otherApprovals = await this.getContractApprovals(approval.ContractId);
      for (const other of otherApprovals) {
        if (other.Id !== approvalId && other.Status === ContractApprovalStatus.Pending) {
          await this.sp.web.lists.getByTitle(LISTS.APPROVALS).items.getById(other.Id!).update({
            Status: ContractApprovalStatus.Cancelled
          });
        }
      }

      // Log audit
      await this.logAudit(approval.ContractId, ContractAuditAction.ApprovalRejected, `Rejected: ${reason}`, 'Approval');
    } catch (error) {
      logger.error('ContractManagementService', `Error rejecting contract:`, error);
      throw error;
    }
  }

  /**
   * Send contract for signature
   */
  public async sendForSignature(contractId: number, signatories: number[]): Promise<void> {
    try {
      const contract = await this.getContractById(contractId);

      if (contract.Status !== ContractLifecycleStatus.Approved) {
        throw new Error('Contract must be approved before sending for signature');
      }

      // Create signature requests for each signatory
      for (let i = 0; i < signatories.length; i++) {
        const party = await this.getContractPartyById(signatories[i]);
        await this.createSignatureRequest(contractId, signatories[i], party, i + 1);
      }

      // Update contract status
      await this.updateContract(contractId, {
        Status: ContractLifecycleStatus.PendingSignature,
        SentForSignatureDate: new Date()
      });

      // Log audit
      await this.logAudit(contractId, ContractAuditAction.SignatureRequested, `Sent for signature to ${signatories.length} parties`, 'Signature');
    } catch (error) {
      logger.error('ContractManagementService', `Error sending contract ${contractId} for signature:`, error);
      throw error;
    }
  }

  /**
   * Record a signature
   */
  public async recordSignature(signatureId: number, signedDate: Date): Promise<void> {
    try {
      const signature = await this.getSignatureById(signatureId);

      await this.sp.web.lists.getByTitle(LISTS.SIGNATURES).items.getById(signatureId).update({
        Status: SignatureStatus.Signed,
        SignedDate: signedDate
      });

      // Check if all signatures are complete
      const allSignatures = await this.getContractSignatures(signature.ContractId);
      const pendingSignatures = allSignatures.filter(s =>
        s.Status === SignatureStatus.Pending || s.Status === SignatureStatus.Sent
      );

      if (pendingSignatures.length === 0) {
        // All signed - mark as fully executed
        await this.updateContract(signature.ContractId, {
          Status: ContractLifecycleStatus.FullyExecuted,
          SignedDate: new Date(),
          FullyExecutedDate: new Date()
        });

        // After a delay, move to Active status
        await this.updateContract(signature.ContractId, {
          Status: ContractLifecycleStatus.Active
        });
      } else {
        // Update to partially signed
        await this.updateContract(signature.ContractId, {
          Status: ContractLifecycleStatus.PartiallySigned
        });
      }

      // Log audit
      await this.logAudit(signature.ContractId, ContractAuditAction.SignatureReceived, 'Signature received', 'Signature');
    } catch (error) {
      logger.error('ContractManagementService', `Error recording signature:`, error);
      throw error;
    }
  }

  /**
   * Terminate a contract
   */
  public async terminateContract(contractId: number, terminationType: string, reason: string): Promise<void> {
    try {
      const contract = await this.getContractById(contractId);

      if (contract.Status !== ContractLifecycleStatus.Active) {
        throw new Error('Only active contracts can be terminated');
      }

      await this.updateContract(contractId, {
        Status: ContractLifecycleStatus.Terminated,
        TerminationDate: new Date(),
        Notes: contract.Notes
          ? `${contract.Notes}\n\nTermination (${terminationType}): ${reason}`
          : `Termination (${terminationType}): ${reason}`
      });

      // Log audit
      await this.logAudit(contractId, ContractAuditAction.Terminated, `Contract terminated: ${reason}`, 'Contract');
    } catch (error) {
      logger.error('ContractManagementService', `Error terminating contract ${contractId}:`, error);
      throw error;
    }
  }

  /**
   * Renew a contract
   */
  public async renewContract(contractId: number, newEndDate: Date, newValue?: number): Promise<IContractRecord> {
    try {
      const contract = await this.getContractById(contractId);

      // Create new contract based on existing
      const renewedContract = await this.createContract({
        Title: `${contract.Title} (Renewed ${new Date().getFullYear()})`,
        Category: contract.Category,
        Priority: contract.Priority,
        Industry: contract.Industry,
        Description: contract.Description,
        EffectiveDate: contract.ExpirationDate, // Start when old one ends
        ExpirationDate: newEndDate,
        RenewalType: contract.RenewalType,
        RenewalTermMonths: contract.RenewalTermMonths,
        RenewalNotificationDays: contract.RenewalNotificationDays,
        TerminationNoticeDays: contract.TerminationNoticeDays,
        ValueType: contract.ValueType,
        TotalValue: newValue || contract.TotalValue,
        AnnualValue: newValue || contract.AnnualValue,
        Currency: contract.Currency,
        PaymentTerms: contract.PaymentTerms,
        OwnerId: contract.OwnerId,
        Department: contract.Department,
        BusinessUnit: contract.BusinessUnit,
        PrimaryCounterpartyId: contract.PrimaryCounterpartyId,
        CounterpartyName: contract.CounterpartyName,
        CounterpartyContact: contract.CounterpartyContact,
        CounterpartyEmail: contract.CounterpartyEmail,
        ParentContractId: contractId,
        AmendmentReason: 'Renewal',
        Notes: `Renewed from ${contract.ContractNumber}`
      });

      // Update original contract
      await this.updateContract(contractId, {
        Status: ContractLifecycleStatus.Renewed,
        Notes: contract.Notes
          ? `${contract.Notes}\n\nRenewed to: ${renewedContract.ContractNumber}`
          : `Renewed to: ${renewedContract.ContractNumber}`
      });

      // Log audit
      await this.logAudit(contractId, ContractAuditAction.Renewed, `Contract renewed to ${renewedContract.ContractNumber}`, 'Contract');

      return renewedContract;
    } catch (error) {
      logger.error('ContractManagementService', `Error renewing contract ${contractId}:`, error);
      throw error;
    }
  }

  // ==================== CLAUSE BANK ====================

  /**
   * Get clauses with optional filtering
   */
  public async getClauses(filter?: IClauseFilter): Promise<IContractClause[]> {
    try {
      let query = this.sp.web.lists.getByTitle(LISTS.CLAUSES).items
        .select(
          'Id', 'Title', 'ClauseCode', 'ClauseName', 'Category', 'SubCategory', 'Industry',
          'ClauseContent', 'PlainLanguageSummary', 'FallbackContent', 'ShortFormContent',
          'RiskLevel', 'Negotiability', 'IsActive', 'IsMandatory', 'IsDefault',
          'RegulatoryRequirement', 'JurisdictionApplicability', 'Variables',
          'RelatedClauseIds', 'ConflictingClauseIds', 'RequiresClauseIds',
          'Version', 'EffectiveDate', 'RetiredDate', 'ReplacedByClauseId',
          'Author', 'LegalReviewDate', 'LegalApprovedById', 'UsageCount',
          'Tags', 'Notes', 'Created', 'Modified'
        );

      // Apply filters
      const filters = this.buildClauseFilters(filter);
      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.orderBy('Category', true).orderBy('ClauseCode', true).top(5000)();
      return items.map(this.mapClauseFromSP);
    } catch (error) {
      logger.error('ContractManagementService', 'Error getting clauses:', error);
      throw error;
    }
  }

  /**
   * Get a single clause by ID
   */
  public async getClauseById(id: number): Promise<IContractClause> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(LISTS.CLAUSES).items
        .getById(validId)
        .select(
          'Id', 'Title', 'ClauseCode', 'ClauseName', 'Category', 'SubCategory', 'Industry',
          'ClauseContent', 'PlainLanguageSummary', 'FallbackContent', 'ShortFormContent',
          'RiskLevel', 'Negotiability', 'IsActive', 'IsMandatory', 'IsDefault',
          'RegulatoryRequirement', 'JurisdictionApplicability', 'Variables',
          'RelatedClauseIds', 'ConflictingClauseIds', 'RequiresClauseIds',
          'Version', 'EffectiveDate', 'RetiredDate', 'ReplacedByClauseId',
          'Author', 'LegalReviewDate', 'LegalApprovedById', 'UsageCount',
          'Tags', 'Notes', 'Created', 'Modified'
        )();

      return this.mapClauseFromSP(item);
    } catch (error) {
      logger.error('ContractManagementService', `Error getting clause ${id}:`, error);
      throw error;
    }
  }

  /**
   * Add a clause instance to a contract
   */
  public async addClauseToContract(
    contractId: number,
    clauseId: number,
    sectionNumber: string,
    displayOrder: number,
    variableValues?: Record<string, string>
  ): Promise<IContractClauseInstance> {
    try {
      const clause = await this.getClauseById(clauseId);

      // Apply variable substitutions
      let content = clause.ClauseContent;
      if (variableValues) {
        for (const [key, value] of Object.entries(variableValues)) {
          content = content.replace(new RegExp(`{{${key}}}`, 'g'), value);
        }
      }

      const itemData = {
        Title: clause.ClauseName,
        ContractId: contractId,
        ClauseId: clauseId,
        SectionNumber: sectionNumber,
        DisplayOrder: displayOrder,
        ClauseContent: content,
        IsModified: !!variableValues,
        ModificationNotes: variableValues ? 'Variables substituted' : null,
        Status: 'Draft',
        VariableValues: variableValues ? JSON.stringify(variableValues) : null,
        AddedById: this.currentUserId,
        IsReviewed: false
      };

      const result = await this.sp.web.lists.getByTitle(LISTS.CLAUSE_INSTANCES).items.add(itemData);

      // Update clause usage count
      await this.sp.web.lists.getByTitle(LISTS.CLAUSES).items.getById(clauseId).update({
        UsageCount: (clause.UsageCount || 0) + 1
      });

      // Log audit
      await this.logAudit(contractId, ContractAuditAction.ClauseAdded, `Added clause: ${clause.ClauseName}`, 'Clause');

      return this.getClauseInstanceById(result.data.Id);
    } catch (error) {
      logger.error('ContractManagementService', `Error adding clause to contract:`, error);
      throw error;
    }
  }

  /**
   * Get clause instances for a contract
   */
  public async getContractClauses(contractId: number): Promise<IContractClauseInstance[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(LISTS.CLAUSE_INSTANCES).items
        .select(
          'Id', 'Title', 'ContractId', 'ClauseId', 'SectionNumber', 'DisplayOrder',
          'ClauseContent', 'IsModified', 'ModificationNotes', 'Status', 'NegotiationNotes',
          'VariableValues', 'AddedById', 'AddedBy/Title', 'ModifiedById', 'ModifiedBy/Title',
          'IsReviewed', 'ReviewedById', 'ReviewedBy/Title', 'ReviewedDate', 'ReviewComments',
          'Created', 'Modified'
        )
        .expand('AddedBy', 'ModifiedBy', 'ReviewedBy')
        .filter(`ContractId eq ${contractId}`)
        .orderBy('DisplayOrder', true)();

      return items.map(this.mapClauseInstanceFromSP);
    } catch (error) {
      logger.error('ContractManagementService', `Error getting contract clauses:`, error);
      throw error;
    }
  }

  // ==================== OBLIGATIONS ====================

  /**
   * Get obligations with optional filtering
   */
  public async getObligations(filter?: IObligationFilter): Promise<IContractObligation[]> {
    try {
      console.log('[ContractManagementService] getObligations - Loading from list:', LISTS.OBLIGATIONS);

      // Simplified query without Person field expands (list may not have these columns)
      let query = this.sp.web.lists.getByTitle(LISTS.OBLIGATIONS).items
        .select(
          'Id', 'Title', 'ContractId', 'ObligationType', 'Description', 'ClauseReference',
          'ResponsibleParty', 'AssigneeId',
          'DueDate', 'Frequency', 'RecurrencePattern', 'NextOccurrence', 'EndDate',
          'Status', 'CompletedDate', 'CompletedById', 'CompletionNotes',
          'ReminderDays', 'RemindersSent', 'LastReminderDate',
          'Amount', 'Currency', 'EvidenceRequired', 'EvidenceDocumentUrl',
          'HasPenalty', 'PenaltyDescription', 'PenaltyAmount', 'Priority',
          'Notes', 'Created', 'Modified'
        );

      // Apply filters
      const filters = this.buildObligationFilters(filter);
      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.orderBy('DueDate', true).top(5000)();
      console.log(`[ContractManagementService] getObligations - Retrieved ${items.length} obligations`);
      return items.map(this.mapObligationFromSP);
    } catch (error: any) {
      console.error('[ContractManagementService] getObligations - Error:', error?.message || error);
      // Return empty array instead of throwing to allow dashboard to load
      return [];
    }
  }

  /**
   * Create an obligation
   */
  public async createObligation(obligation: Partial<IContractObligation>): Promise<IContractObligation> {
    try {
      if (!obligation.ContractId || !obligation.Description || !obligation.DueDate) {
        throw new Error('ContractId, Description, and DueDate are required');
      }

      const itemData = {
        Title: obligation.Title || obligation.Description.substring(0, 255),
        ContractId: obligation.ContractId,
        ObligationType: obligation.ObligationType,
        Description: obligation.Description,
        ClauseReference: obligation.ClauseReference,
        ResponsibleParty: obligation.ResponsibleParty || 'Internal',
        AssigneeId: obligation.AssigneeId,
        DueDate: obligation.DueDate,
        Frequency: obligation.Frequency || ObligationFrequency.OneTime,
        RecurrencePattern: obligation.RecurrencePattern,
        NextOccurrence: obligation.NextOccurrence,
        EndDate: obligation.EndDate,
        Status: ObligationStatus.Upcoming,
        ReminderDays: obligation.ReminderDays || 7,
        Amount: obligation.Amount,
        Currency: obligation.Currency,
        EvidenceRequired: obligation.EvidenceRequired || false,
        HasPenalty: obligation.HasPenalty || false,
        PenaltyDescription: obligation.PenaltyDescription,
        PenaltyAmount: obligation.PenaltyAmount,
        Priority: obligation.Priority || ContractPriority.Medium,
        Notes: obligation.Notes
      };

      const result = await this.sp.web.lists.getByTitle(LISTS.OBLIGATIONS).items.add(itemData);

      // Log audit
      await this.logAudit(obligation.ContractId, ContractAuditAction.ObligationAdded, `Obligation added: ${obligation.Description}`, 'Obligation');

      return this.getObligationById(result.data.Id);
    } catch (error) {
      logger.error('ContractManagementService', 'Error creating obligation:', error);
      throw error;
    }
  }

  /**
   * Complete an obligation
   */
  public async completeObligation(obligationId: number, completionNotes?: string, evidenceUrl?: string): Promise<void> {
    try {
      const obligation = await this.getObligationById(obligationId);

      const updateData: Record<string, unknown> = {
        Status: ObligationStatus.Completed,
        CompletedDate: new Date(),
        CompletedById: this.currentUserId,
        CompletionNotes: completionNotes
      };

      if (evidenceUrl) {
        updateData.EvidenceDocumentUrl = evidenceUrl;
      }

      await this.sp.web.lists.getByTitle(LISTS.OBLIGATIONS).items.getById(obligationId).update(updateData);

      // If recurring, create next occurrence
      if (obligation.Frequency !== ObligationFrequency.OneTime && obligation.EndDate) {
        const nextDueDate = this.calculateNextOccurrence(obligation.DueDate, obligation.Frequency);
        if (nextDueDate <= new Date(obligation.EndDate)) {
          await this.createObligation({
            ContractId: obligation.ContractId,
            Title: obligation.Title,
            ObligationType: obligation.ObligationType,
            Description: obligation.Description,
            ClauseReference: obligation.ClauseReference,
            ResponsibleParty: obligation.ResponsibleParty,
            AssigneeId: obligation.AssigneeId,
            DueDate: nextDueDate,
            Frequency: obligation.Frequency,
            RecurrencePattern: obligation.RecurrencePattern,
            EndDate: obligation.EndDate,
            ReminderDays: obligation.ReminderDays,
            Amount: obligation.Amount,
            Currency: obligation.Currency,
            EvidenceRequired: obligation.EvidenceRequired,
            HasPenalty: obligation.HasPenalty,
            PenaltyDescription: obligation.PenaltyDescription,
            PenaltyAmount: obligation.PenaltyAmount,
            Priority: obligation.Priority,
            Notes: obligation.Notes
          });
        }
      }

      // Log audit
      await this.logAudit(obligation.ContractId, ContractAuditAction.ObligationCompleted, `Obligation completed: ${obligation.Title}`, 'Obligation');
    } catch (error) {
      logger.error('ContractManagementService', `Error completing obligation ${obligationId}:`, error);
      throw error;
    }
  }

  // ==================== DASHBOARD & STATISTICS ====================

  /**
   * Get contract statistics
   */
  public async getStatistics(): Promise<IContractStatistics> {
    try {
      console.log('[ContractManagementService] getStatistics - Loading statistics data');
      console.log('[ContractManagementService] SP object:', this.sp ? 'initialized' : 'NOT initialized');

      const contracts = await this.getContracts();
      console.log(`[ContractManagementService] getStatistics - Loaded ${contracts.length} contracts`);
      const obligations = await this.getObligations();
      console.log(`[ContractManagementService] getStatistics - Loaded ${obligations.length} obligations`);

      const today = new Date();
      const in30Days = new Date();
      in30Days.setDate(today.getDate() + 30);
      const in60Days = new Date();
      in60Days.setDate(today.getDate() + 60);
      const in90Days = new Date();
      in90Days.setDate(today.getDate() + 90);

      const stats: IContractStatistics = {
        totalContracts: contracts.length,
        activeContracts: 0,
        draftContracts: 0,
        pendingApproval: 0,
        pendingSignature: 0,
        expiredContracts: 0,
        expiring30Days: 0,
        expiring60Days: 0,
        expiring90Days: 0,
        contractsByStatus: {},
        contractsByCategory: {},
        contractsByRisk: {},
        totalContractValue: 0,
        totalAnnualValue: 0,
        avgContractValue: 0,
        valueByCategory: {},
        valueByCurrency: {},
        avgCycleTime: 0,
        avgApprovalTime: 0,
        avgSignatureTime: 0,
        totalObligations: obligations.length,
        overdueObligations: 0,
        upcomingObligations: 0,
        contractsCreatedThisMonth: 0,
        contractsExecutedThisMonth: 0,
        renewalRate: 0
      };

      const thisMonth = new Date();
      thisMonth.setDate(1);
      thisMonth.setHours(0, 0, 0, 0);

      let totalValue = 0;
      let activeCount = 0;

      for (const contract of contracts) {
        // Count by status
        stats.contractsByStatus[contract.Status] = (stats.contractsByStatus[contract.Status] || 0) + 1;

        // Count by category
        stats.contractsByCategory[contract.Category] = (stats.contractsByCategory[contract.Category] || 0) + 1;

        // Count by risk
        if (contract.RiskLevel) {
          stats.contractsByRisk[contract.RiskLevel] = (stats.contractsByRisk[contract.RiskLevel] || 0) + 1;
        }

        // Status-based counts
        switch (contract.Status) {
          case ContractLifecycleStatus.Draft:
            stats.draftContracts++;
            break;
          case ContractLifecycleStatus.PendingApproval:
            stats.pendingApproval++;
            break;
          case ContractLifecycleStatus.PendingSignature:
          case ContractLifecycleStatus.PartiallySigned:
            stats.pendingSignature++;
            break;
          case ContractLifecycleStatus.Active:
            stats.activeContracts++;
            activeCount++;
            totalValue += contract.TotalValue || 0;
            stats.totalAnnualValue += contract.AnnualValue || 0;

            // Value by category
            stats.valueByCategory[contract.Category] = (stats.valueByCategory[contract.Category] || 0) + (contract.TotalValue || 0);

            // Value by currency
            if (contract.Currency) {
              stats.valueByCurrency[contract.Currency] = (stats.valueByCurrency[contract.Currency] || 0) + (contract.TotalValue || 0);
            }

            // Expiry checks
            if (contract.ExpirationDate) {
              const expDate = new Date(contract.ExpirationDate);
              if (expDate < today) {
                stats.expiredContracts++;
              } else if (expDate <= in30Days) {
                stats.expiring30Days++;
                stats.expiring60Days++;
                stats.expiring90Days++;
              } else if (expDate <= in60Days) {
                stats.expiring60Days++;
                stats.expiring90Days++;
              } else if (expDate <= in90Days) {
                stats.expiring90Days++;
              }
            }
            break;
          case ContractLifecycleStatus.Expired:
            stats.expiredContracts++;
            break;
        }

        // This month metrics
        if (contract.Created && new Date(contract.Created) >= thisMonth) {
          stats.contractsCreatedThisMonth++;
        }
        if (contract.FullyExecutedDate && new Date(contract.FullyExecutedDate) >= thisMonth) {
          stats.contractsExecutedThisMonth++;
        }
      }

      stats.totalContractValue = totalValue;
      stats.avgContractValue = activeCount > 0 ? totalValue / activeCount : 0;

      // Obligation metrics
      for (const obligation of obligations) {
        if (obligation.Status === ObligationStatus.Overdue) {
          stats.overdueObligations++;
        } else if (obligation.Status === ObligationStatus.Upcoming || obligation.Status === ObligationStatus.Due) {
          const dueDate = new Date(obligation.DueDate);
          if (dueDate >= today && dueDate <= in30Days) {
            stats.upcomingObligations++;
          }
        }
      }

      return stats;
    } catch (error) {
      logger.error('ContractManagementService', 'Error getting statistics:', error);
      throw error;
    }
  }

  /**
   * Get dashboard data
   */
  public async getDashboard(): Promise<IContractDashboard> {
    try {
      const [statistics, contracts, approvals, signatures, obligations, auditLog] = await Promise.all([
        this.getStatistics(),
        this.getContracts(),
        this.getApprovals({ status: [ContractApprovalStatus.Pending] }),
        this.getSignatures({ status: [SignatureStatus.Pending, SignatureStatus.Sent] }),
        this.getObligations({ isOverdue: false }),
        this.getAuditLog(undefined, 20)
      ]);

      const today = new Date();
      const in90Days = new Date();
      in90Days.setDate(today.getDate() + 90);

      // Expiring contracts
      const expiringContracts = contracts.filter(c =>
        c.Status === ContractLifecycleStatus.Active &&
        c.ExpirationDate &&
        new Date(c.ExpirationDate) <= in90Days
      ).slice(0, 10);

      // My contracts (owned by current user)
      const myContracts = contracts.filter(c => c.OwnerId === this.currentUserId).slice(0, 10);

      // Upcoming obligations (next 30 days)
      const in30Days = new Date();
      in30Days.setDate(today.getDate() + 30);
      const upcomingObligations = obligations.filter(o =>
        (o.Status === ObligationStatus.Upcoming || o.Status === ObligationStatus.Due) &&
        new Date(o.DueDate) <= in30Days
      ).slice(0, 10);

      // Generate alerts
      const alerts = this.generateAlerts(contracts, obligations);

      // Expiry timeline (next 12 months)
      const expiryTimeline = this.generateExpiryTimeline(contracts);

      // Value by department
      const valueByDepartment = this.calculateValueByDepartment(contracts);

      return {
        statistics,
        expiringContracts,
        pendingApprovals: approvals.slice(0, 10),
        pendingSignatures: signatures.slice(0, 10),
        upcomingObligations,
        recentActivity: auditLog,
        myContracts,
        alerts,
        expiryTimeline,
        valueByDepartment
      };
    } catch (error) {
      logger.error('ContractManagementService', 'Error getting dashboard:', error);
      throw error;
    }
  }

  // ==================== AUDIT LOG ====================

  /**
   * Log an audit entry
   */
  public async logAudit(
    contractId: number,
    action: ContractAuditAction,
    description: string,
    category: string,
    previousValue?: string,
    newValue?: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(LISTS.AUDIT_LOG).items.add({
        Title: `${action} - Contract ${contractId}`,
        ContractId: contractId,
        Action: action,
        ActionCategory: category,
        ActionDescription: description,
        PreviousValue: previousValue,
        NewValue: newValue,
        ActionById: this.currentUserId,
        ActionDate: new Date(),
        IsSystemAction: false,
        Severity: this.getAuditSeverity(action)
      });
    } catch (error) {
      // Don't throw on audit failures - just log
      logger.error('ContractManagementService', 'Error logging audit:', error);
    }
  }

  /**
   * Get audit log entries
   */
  public async getAuditLog(contractId?: number, limit: number = 100): Promise<IContractAuditLog[]> {
    try {
      console.log('[ContractManagementService] getAuditLog - Loading from list:', LISTS.AUDIT_LOG);

      // Simplified query without Person field expands (causes 400 errors)
      let query = this.sp.web.lists.getByTitle(LISTS.AUDIT_LOG).items
        .select(
          'Id', 'Title', 'ContractId', 'Action', 'ActionCategory', 'ActionDescription',
          'PreviousValue', 'NewValue', 'ChangeDetails', 'RelatedEntityType', 'RelatedEntityId',
          'ActionById', 'ActionDate', 'IPAddress', 'UserAgent', 'SessionId',
          'IsSystemAction', 'Severity', 'Notes', 'Created'
        );

      if (contractId) {
        query = query.filter(`ContractId eq ${contractId}`);
      }

      const items = await query.orderBy('ActionDate', false).top(limit)();
      console.log(`[ContractManagementService] getAuditLog - Retrieved ${items.length} entries`);
      return items.map(this.mapAuditLogFromSP);
    } catch (error: any) {
      console.error('[ContractManagementService] getAuditLog - Error:', error?.message || error);
      // Return empty array instead of throwing to prevent UI crashes
      return [];
    }
  }

  // ==================== TEMPLATES ====================

  /**
   * Get contract templates
   */
  public async getTemplates(category?: ContractCategory): Promise<IContractTemplate[]> {
    try {
      console.log('[ContractManagementService] getTemplates - Loading from list:', LISTS.TEMPLATES);
      console.log('[ContractManagementService] getTemplates - Category filter:', category || 'none');
      console.log('[ContractManagementService] SP object:', this.sp ? 'initialized' : 'NOT initialized');

      let query = this.sp.web.lists.getByTitle(LISTS.TEMPLATES).items
        .select(
          'Id', 'Title', 'TemplateCode', 'TemplateName', 'Category', 'Industry',
          'Description', 'UsageGuidance', 'TemplateDocumentUrl', 'DefaultClauses', 'MandatoryClauses',
          'DefaultDurationMonths', 'DefaultRenewalType', 'DefaultNotificationDays',
          'DefaultCurrency', 'DefaultPaymentTerms', 'DefaultApproverIds', 'ApprovalThresholds',
          'Variables', 'IsActive', 'IsPublished', 'Version', 'RequiresLegalReview',
          'ComplianceChecklist', 'CreatedById', 'CreatedBy/Title', 'LastModifiedById', 'LastModifiedBy/Title',
          'UsageCount', 'Tags', 'Notes', 'Created', 'Modified'
        )
        .expand('CreatedBy', 'LastModifiedBy')
        .filter('IsActive eq true and IsPublished eq true');

      if (category) {
        query = query.filter(`Category eq '${category}'`);
      }

      const items = await query.orderBy('TemplateName', true).top(500)();
      console.log(`[ContractManagementService] getTemplates - Retrieved ${items.length} templates from ${LISTS.TEMPLATES}`);
      return items.map(this.mapTemplateFromSP);
    } catch (error: any) {
      console.error(`[ContractManagementService] getTemplates - Error loading templates from ${LISTS.TEMPLATES}:`, error?.message || error);
      logger.error('ContractManagementService', 'Error getting templates:', error);
      throw error;
    }
  }

  /**
   * Create contract from template
   */
  public async createFromTemplate(
    templateId: number,
    contractData: Partial<IContractRecord>
  ): Promise<IContractRecord> {
    try {
      const template = await this.getTemplateById(templateId);

      // Merge template defaults with provided data
      const mergedData: Partial<IContractRecord> = {
        ...contractData,
        Category: contractData.Category || template.Category,
        RenewalType: contractData.RenewalType || template.DefaultRenewalType,
        RenewalNotificationDays: contractData.RenewalNotificationDays || template.DefaultNotificationDays,
        Currency: contractData.Currency || template.DefaultCurrency,
        PaymentTerms: contractData.PaymentTerms || template.DefaultPaymentTerms,
        RequiresLegalReview: template.RequiresLegalReview,
        TemplateId: templateId,
        TemplateName: template.TemplateName
      };

      // Calculate expiration date if not provided
      if (!mergedData.ExpirationDate && mergedData.EffectiveDate && template.DefaultDurationMonths) {
        const expDate = new Date(mergedData.EffectiveDate);
        expDate.setMonth(expDate.getMonth() + template.DefaultDurationMonths);
        mergedData.ExpirationDate = expDate;
      }

      // Create the contract
      const contract = await this.createContract(mergedData);

      // Add mandatory clauses
      if (template.MandatoryClauses) {
        const mandatoryClauseIds = JSON.parse(template.MandatoryClauses) as number[];
        for (let i = 0; i < mandatoryClauseIds.length; i++) {
          await this.addClauseToContract(
            contract.Id!,
            mandatoryClauseIds[i],
            `${i + 1}`,
            i + 1
          );
        }
      }

      // Update template usage count
      await this.sp.web.lists.getByTitle(LISTS.TEMPLATES).items.getById(templateId).update({
        UsageCount: (template.UsageCount || 0) + 1
      });

      return this.getContractById(contract.Id!);
    } catch (error) {
      logger.error('ContractManagementService', `Error creating contract from template ${templateId}:`, error);
      throw error;
    }
  }

  // ==================== PRIVATE HELPER METHODS ====================

  private async generateContractNumber(): Promise<string> {
    try {
      const year = new Date().getFullYear();
      const prefix = `CON-${year}-`;

      const items = await this.sp.web.lists.getByTitle(LISTS.CONTRACTS).items
        .select('ContractNumber')
        .filter(`substringof('${prefix}', ContractNumber)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].ContractNumber) {
        const match = items[0].ContractNumber.match(/CON-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(4, '0')}`;
    } catch {
      return `CON-${Date.now()}`;
    }
  }

  private buildContractFilters(filter?: IContractFilter): string[] {
    const filters: string[] = [];
    if (!filter) return filters;

    if (filter.searchTerm) {
      const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
      filters.push(`(substringof('${term}', Title) or substringof('${term}', ContractNumber) or substringof('${term}', CounterpartyName))`);
    }

    if (filter.status && filter.status.length > 0) {
      const statusFilters = filter.status.map(s => `Status eq '${s}'`);
      filters.push(`(${statusFilters.join(' or ')})`);
    }

    if (filter.category && filter.category.length > 0) {
      const catFilters = filter.category.map(c => `Category eq '${c}'`);
      filters.push(`(${catFilters.join(' or ')})`);
    }

    if (filter.priority && filter.priority.length > 0) {
      const prioFilters = filter.priority.map(p => `Priority eq '${p}'`);
      filters.push(`(${prioFilters.join(' or ')})`);
    }

    if (filter.riskLevel && filter.riskLevel.length > 0) {
      const riskFilters = filter.riskLevel.map(r => `RiskLevel eq '${r}'`);
      filters.push(`(${riskFilters.join(' or ')})`);
    }

    if (filter.ownerId !== undefined) {
      filters.push(`OwnerId eq ${filter.ownerId}`);
    }

    if (filter.department) {
      filters.push(`Department eq '${ValidationUtils.sanitizeForOData(filter.department)}'`);
    }

    if (filter.counterpartyId !== undefined) {
      filters.push(`PrimaryCounterpartyId eq ${filter.counterpartyId}`);
    }

    if (filter.counterpartyName) {
      filters.push(`substringof('${ValidationUtils.sanitizeForOData(filter.counterpartyName)}', CounterpartyName)`);
    }

    if (filter.expiringWithinDays !== undefined) {
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + filter.expiringWithinDays);
      filters.push(`ExpirationDate le datetime'${futureDate.toISOString()}'`);
      filters.push(`Status eq '${ContractLifecycleStatus.Active}'`);
    }

    if (filter.minValue !== undefined) {
      filters.push(`TotalValue ge ${filter.minValue}`);
    }

    if (filter.maxValue !== undefined) {
      filters.push(`TotalValue le ${filter.maxValue}`);
    }

    if (filter.isAmendment !== undefined) {
      filters.push(`IsAmendment eq ${filter.isAmendment}`);
    }

    if (filter.parentContractId !== undefined) {
      filters.push(`ParentContractId eq ${filter.parentContractId}`);
    }

    return filters;
  }

  private buildClauseFilters(filter?: IClauseFilter): string[] {
    const filters: string[] = [];
    if (!filter) return filters;

    if (filter.searchTerm) {
      const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
      filters.push(`(substringof('${term}', ClauseName) or substringof('${term}', ClauseCode) or substringof('${term}', ClauseContent))`);
    }

    if (filter.category && filter.category.length > 0) {
      const catFilters = filter.category.map(c => `Category eq '${c}'`);
      filters.push(`(${catFilters.join(' or ')})`);
    }

    if (filter.riskLevel && filter.riskLevel.length > 0) {
      const riskFilters = filter.riskLevel.map(r => `RiskLevel eq '${r}'`);
      filters.push(`(${riskFilters.join(' or ')})`);
    }

    if (filter.isActive !== undefined) {
      filters.push(`IsActive eq ${filter.isActive}`);
    }

    if (filter.isMandatory !== undefined) {
      filters.push(`IsMandatory eq ${filter.isMandatory}`);
    }

    return filters;
  }

  private buildObligationFilters(filter?: IObligationFilter): string[] {
    const filters: string[] = [];
    if (!filter) return filters;

    if (filter.contractId !== undefined) {
      filters.push(`ContractId eq ${filter.contractId}`);
    }

    if (filter.type && filter.type.length > 0) {
      const typeFilters = filter.type.map(t => `ObligationType eq '${t}'`);
      filters.push(`(${typeFilters.join(' or ')})`);
    }

    if (filter.status && filter.status.length > 0) {
      const statusFilters = filter.status.map(s => `Status eq '${s}'`);
      filters.push(`(${statusFilters.join(' or ')})`);
    }

    if (filter.assigneeId !== undefined) {
      filters.push(`AssigneeId eq ${filter.assigneeId}`);
    }

    if (filter.dueDateFrom) {
      filters.push(`DueDate ge datetime'${filter.dueDateFrom.toISOString()}'`);
    }

    if (filter.dueDateTo) {
      filters.push(`DueDate le datetime'${filter.dueDateTo.toISOString()}'`);
    }

    if (filter.isOverdue) {
      filters.push(`DueDate lt datetime'${new Date().toISOString()}'`);
      filters.push(`Status ne '${ObligationStatus.Completed}'`);
    }

    return filters;
  }

  private generateAlerts(contracts: IContractRecord[], obligations: IContractObligation[]): IContractAlert[] {
    const alerts: IContractAlert[] = [];
    const today = new Date();
    const in30Days = new Date();
    in30Days.setDate(today.getDate() + 30);

    // Expiring contracts
    for (const contract of contracts) {
      if (contract.Status === ContractLifecycleStatus.Active && contract.ExpirationDate) {
        const expDate = new Date(contract.ExpirationDate);
        if (expDate <= in30Days) {
          const daysUntil = Math.ceil((expDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
          alerts.push({
            contractId: contract.Id!,
            contractNumber: contract.ContractNumber,
            contractTitle: contract.Title,
            alertType: 'Expiry',
            severity: daysUntil <= 7 ? 'Critical' : daysUntil <= 14 ? 'Warning' : 'Info',
            message: `Contract expires in ${daysUntil} days`,
            dueDate: expDate
          });
        }
      }
    }

    // Overdue obligations
    for (const obligation of obligations) {
      if (obligation.Status !== ObligationStatus.Completed) {
        const dueDate = new Date(obligation.DueDate);
        if (dueDate < today) {
          alerts.push({
            contractId: obligation.ContractId,
            contractNumber: '',
            contractTitle: obligation.Title,
            alertType: 'Obligation',
            severity: 'Critical',
            message: `Obligation overdue since ${dueDate.toLocaleDateString()}`,
            dueDate
          });
        } else if (dueDate <= in30Days) {
          const daysUntil = Math.ceil((dueDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
          alerts.push({
            contractId: obligation.ContractId,
            contractNumber: '',
            contractTitle: obligation.Title,
            alertType: 'Obligation',
            severity: daysUntil <= 7 ? 'Warning' : 'Info',
            message: `Obligation due in ${daysUntil} days`,
            dueDate
          });
        }
      }
    }

    return alerts.sort((a, b) => {
      const severityOrder = { Critical: 0, Warning: 1, Info: 2 };
      return severityOrder[a.severity] - severityOrder[b.severity];
    }).slice(0, 20);
  }

  private generateExpiryTimeline(contracts: IContractRecord[]): IExpiryTimelineItem[] {
    const timeline: IExpiryTimelineItem[] = [];
    const today = new Date();

    // Next 12 months
    for (let i = 0; i < 12; i++) {
      const month = new Date(today.getFullYear(), today.getMonth() + i, 1);
      const monthEnd = new Date(today.getFullYear(), today.getMonth() + i + 1, 0);
      const monthStr = month.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });

      const expiringThisMonth = contracts.filter(c =>
        c.Status === ContractLifecycleStatus.Active &&
        c.ExpirationDate &&
        new Date(c.ExpirationDate) >= month &&
        new Date(c.ExpirationDate) <= monthEnd
      );

      timeline.push({
        month: monthStr,
        count: expiringThisMonth.length,
        value: expiringThisMonth.reduce((sum, c) => sum + (c.TotalValue || 0), 0),
        contracts: expiringThisMonth.map(c => ({
          id: c.Id!,
          title: c.Title,
          value: c.TotalValue || 0
        }))
      });
    }

    return timeline;
  }

  private calculateValueByDepartment(contracts: IContractRecord[]): { department: string; value: number }[] {
    const byDept: Record<string, number> = {};

    for (const contract of contracts) {
      if (contract.Status === ContractLifecycleStatus.Active && contract.Department) {
        byDept[contract.Department] = (byDept[contract.Department] || 0) + (contract.TotalValue || 0);
      }
    }

    return Object.entries(byDept)
      .map(([department, value]) => ({ department, value }))
      .sort((a, b) => b.value - a.value);
  }

  private calculateNextOccurrence(currentDueDate: Date, frequency: ObligationFrequency): Date {
    const next = new Date(currentDueDate);

    switch (frequency) {
      case ObligationFrequency.Daily:
        next.setDate(next.getDate() + 1);
        break;
      case ObligationFrequency.Weekly:
        next.setDate(next.getDate() + 7);
        break;
      case ObligationFrequency.BiWeekly:
        next.setDate(next.getDate() + 14);
        break;
      case ObligationFrequency.Monthly:
        next.setMonth(next.getMonth() + 1);
        break;
      case ObligationFrequency.Quarterly:
        next.setMonth(next.getMonth() + 3);
        break;
      case ObligationFrequency.SemiAnnual:
        next.setMonth(next.getMonth() + 6);
        break;
      case ObligationFrequency.Annual:
        next.setFullYear(next.getFullYear() + 1);
        break;
    }

    return next;
  }

  private getAuditSeverity(action: ContractAuditAction): 'Info' | 'Warning' | 'Critical' {
    const criticalActions = [
      ContractAuditAction.Terminated,
      ContractAuditAction.SignatureDeclined,
      ContractAuditAction.ApprovalRejected
    ];

    const warningActions = [
      ContractAuditAction.StatusChanged,
      ContractAuditAction.AmendmentCreated,
      ContractAuditAction.ClauseModified
    ];

    if (criticalActions.includes(action)) return 'Critical';
    if (warningActions.includes(action)) return 'Warning';
    return 'Info';
  }

  private async deleteContractRelatedRecords(contractId: number): Promise<void> {
    // Delete related records (parties, clauses, approvals, etc.)
    const lists = [
      LISTS.PARTIES,
      LISTS.VERSIONS,
      LISTS.CLAUSE_INSTANCES,
      LISTS.APPROVALS,
      LISTS.SIGNATURES,
      LISTS.OBLIGATIONS,
      LISTS.COMMENTS,
      LISTS.DOCUMENTS
    ];

    for (const listName of lists) {
      try {
        const items = await this.sp.web.lists.getByTitle(listName).items
          .select('Id')
          .filter(`ContractId eq ${contractId}`)
          .top(5000)();

        for (const item of items) {
          await this.sp.web.lists.getByTitle(listName).items.getById(item.Id).delete();
        }
      } catch {
        // Ignore errors - list may not exist
      }
    }
  }

  // ==================== ADDITIONAL GETTER METHODS ====================

  private async getApprovalById(id: number): Promise<IContractApproval> {
    const item = await this.sp.web.lists.getByTitle(LISTS.APPROVALS).items.getById(id)();
    return this.mapApprovalFromSP(item);
  }

  private async getSignatureById(id: number): Promise<IContractSignature> {
    const item = await this.sp.web.lists.getByTitle(LISTS.SIGNATURES).items.getById(id)();
    return this.mapSignatureFromSP(item);
  }

  private async getObligationById(id: number): Promise<IContractObligation> {
    const item = await this.sp.web.lists.getByTitle(LISTS.OBLIGATIONS).items
      .getById(id)
      .select(
        'Id', 'Title', 'ContractId', 'ObligationType', 'Description', 'ClauseReference',
        'ResponsibleParty', 'AssigneeId', 'DueDate', 'Frequency', 'RecurrencePattern',
        'NextOccurrence', 'EndDate', 'Status', 'CompletedDate', 'CompletedById',
        'CompletionNotes', 'ReminderDays', 'Amount', 'Currency', 'EvidenceRequired',
        'EvidenceDocumentUrl', 'HasPenalty', 'PenaltyDescription', 'PenaltyAmount',
        'Priority', 'Notes', 'Created', 'Modified'
      )();
    return this.mapObligationFromSP(item);
  }

  private async getClauseInstanceById(id: number): Promise<IContractClauseInstance> {
    const item = await this.sp.web.lists.getByTitle(LISTS.CLAUSE_INSTANCES).items.getById(id)();
    return this.mapClauseInstanceFromSP(item);
  }

  private async getContractPartyById(id: number): Promise<IContractParty> {
    const item = await this.sp.web.lists.getByTitle(LISTS.PARTIES).items.getById(id)();
    return this.mapPartyFromSP(item);
  }

  private async getTemplateById(id: number): Promise<IContractTemplate> {
    const item = await this.sp.web.lists.getByTitle(LISTS.TEMPLATES).items.getById(id)();
    return this.mapTemplateFromSP(item);
  }

  private async getContractApprovals(contractId: number): Promise<IContractApproval[]> {
    const items = await this.sp.web.lists.getByTitle(LISTS.APPROVALS).items
      .filter(`ContractId eq ${contractId}`)();
    return items.map(this.mapApprovalFromSP);
  }

  private async getContractSignatures(contractId: number): Promise<IContractSignature[]> {
    const items = await this.sp.web.lists.getByTitle(LISTS.SIGNATURES).items
      .filter(`ContractId eq ${contractId}`)();
    return items.map(this.mapSignatureFromSP);
  }

  private async getApprovals(filter?: { status?: ContractApprovalStatus[] }): Promise<IContractApproval[]> {
    let query = this.sp.web.lists.getByTitle(LISTS.APPROVALS).items;

    if (filter?.status && filter.status.length > 0) {
      const statusFilters = filter.status.map(s => `Status eq '${s}'`);
      query = query.filter(statusFilters.join(' or '));
    }

    const items = await query.orderBy('RequestedDate', false).top(100)();
    return items.map(this.mapApprovalFromSP);
  }

  private async getSignatures(filter?: { status?: SignatureStatus[] }): Promise<IContractSignature[]> {
    let query = this.sp.web.lists.getByTitle(LISTS.SIGNATURES).items;

    if (filter?.status && filter.status.length > 0) {
      const statusFilters = filter.status.map(s => `Status eq '${s}'`);
      query = query.filter(statusFilters.join(' or '));
    }

    const items = await query.orderBy('RequestedDate', false).top(100)();
    return items.map(this.mapSignatureFromSP);
  }

  private async getApplicableApprovalRules(contract: IContractRecord): Promise<IContractApprovalRule[]> {
    const items = await this.sp.web.lists.getByTitle(LISTS.APPROVAL_RULES).items
      .filter(`IsActive eq true`)
      .orderBy('ApprovalOrder', true)();

    return items
      .filter(rule => {
        // Filter by category if specified
        if (rule.ContractCategory && rule.ContractCategory !== contract.Category) {
          return false;
        }
        // Filter by value thresholds
        if (rule.MinValue && (contract.TotalValue || 0) < rule.MinValue) {
          return false;
        }
        if (rule.MaxValue && (contract.TotalValue || 0) > rule.MaxValue) {
          return false;
        }
        // Filter by department if specified
        if (rule.Department && rule.Department !== contract.Department) {
          return false;
        }
        return true;
      })
      .map(this.mapApprovalRuleFromSP);
  }

  private async createApprovalRequest(
    contractId: number,
    approverId: number,
    stage: number,
    comments?: string
  ): Promise<void> {
    await this.sp.web.lists.getByTitle(LISTS.APPROVALS).items.add({
      Title: `Approval Request - Contract ${contractId}`,
      ContractId: contractId,
      ApprovalStage: stage,
      ApproverId: approverId,
      Status: ContractApprovalStatus.Pending,
      RequestedDate: new Date(),
      RequestComments: comments
    });
  }

  private async createSignatureRequest(
    contractId: number,
    partyId: number,
    party: IContractParty,
    order: number
  ): Promise<void> {
    await this.sp.web.lists.getByTitle(LISTS.SIGNATURES).items.add({
      Title: `Signature Request - ${party.ExternalName || 'Internal'}`,
      ContractId: contractId,
      ContractPartyId: partyId,
      Provider: 'Internal',
      Status: SignatureStatus.Pending,
      RequestedDate: new Date(),
      SignerName: party.ExternalName || '',
      SignerEmail: party.ExternalEmail || '',
      SignerTitle: party.ExternalTitle,
      SignerCompany: party.ExternalCompany
    });
  }

  // ==================== MAPPING FUNCTIONS ====================

  private mapContractFromSP(item: Record<string, unknown>): IContractRecord {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractNumber: item.ContractNumber as string,
      ExternalReference: item.ExternalReference as string,
      Category: item.Category as ContractCategory,
      Status: item.Status as ContractLifecycleStatus,
      Priority: item.Priority as ContractPriority,
      RiskLevel: item.RiskLevel as ContractRiskLevel,
      Industry: item.Industry as string,
      Description: item.Description as string,
      ExecutiveSummary: item.ExecutiveSummary as string,
      EffectiveDate: item.EffectiveDate ? new Date(item.EffectiveDate as string) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate as string) : undefined,
      SignedDate: item.SignedDate ? new Date(item.SignedDate as string) : undefined,
      TerminationDate: item.TerminationDate ? new Date(item.TerminationDate as string) : undefined,
      OriginalStartDate: item.OriginalStartDate ? new Date(item.OriginalStartDate as string) : undefined,
      RenewalType: item.RenewalType as ContractRenewalType,
      RenewalTermMonths: item.RenewalTermMonths as number,
      RenewalNotificationDays: item.RenewalNotificationDays as number || 90,
      NextRenewalDate: item.NextRenewalDate ? new Date(item.NextRenewalDate as string) : undefined,
      TerminationNoticeDays: item.TerminationNoticeDays as number || 30,
      ValueType: item.ValueType as ContractValueType,
      TotalValue: item.TotalValue as number,
      AnnualValue: item.AnnualValue as number,
      MonthlyValue: item.MonthlyValue as number,
      Currency: item.Currency as Currency || Currency.GBP,
      PaymentTerms: item.PaymentTerms as PaymentTerms,
      BudgetCode: item.BudgetCode as string,
      CostCenter: item.CostCenter as string,
      OwnerId: item.OwnerId as number,
      SecondaryOwnerIds: item.SecondaryOwnerIds as string,
      Department: item.Department as string,
      BusinessUnit: item.BusinessUnit as string,
      PrimaryCounterpartyId: item.PrimaryCounterpartyId as number,
      CounterpartyName: item.CounterpartyName as string,
      CounterpartyContact: item.CounterpartyContact as string,
      CounterpartyEmail: item.CounterpartyEmail as string,
      DocumentLibraryUrl: item.DocumentLibraryUrl as string,
      CurrentVersionUrl: item.CurrentVersionUrl as string,
      ExecutedDocumentUrl: item.ExecutedDocumentUrl as string,
      Version: item.Version as number || 1,
      IsAmendment: item.IsAmendment as boolean || false,
      ParentContractId: item.ParentContractId as number,
      AmendmentReason: item.AmendmentReason as string,
      TemplateId: item.TemplateId as number,
      TemplateName: item.TemplateName as string,
      ComplianceRequirements: item.ComplianceRequirements as string,
      DataClassification: item.DataClassification as string,
      GDPRApplicable: item.GDPRApplicable as boolean,
      RequiresLegalReview: item.RequiresLegalReview as boolean,
      RiskScore: item.RiskScore as number,
      RiskFactors: item.RiskFactors as string,
      HasSLATerms: item.HasSLATerms as boolean,
      Tags: item.Tags as string,
      Notes: item.Notes as string,
      CurrentApprovalStage: item.CurrentApprovalStage as number,
      TotalApprovalStages: item.TotalApprovalStages as number,
      SubmittedForApprovalDate: item.SubmittedForApprovalDate ? new Date(item.SubmittedForApprovalDate as string) : undefined,
      ApprovedDate: item.ApprovedDate ? new Date(item.ApprovedDate as string) : undefined,
      SentForSignatureDate: item.SentForSignatureDate ? new Date(item.SentForSignatureDate as string) : undefined,
      FullyExecutedDate: item.FullyExecutedDate ? new Date(item.FullyExecutedDate as string) : undefined,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapClauseFromSP(item: Record<string, unknown>): IContractClause {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ClauseCode: item.ClauseCode as string,
      ClauseName: item.ClauseName as string,
      Category: item.Category as ClauseCategory,
      SubCategory: item.SubCategory as string,
      Industry: item.Industry as string,
      ClauseContent: item.ClauseContent as string,
      PlainLanguageSummary: item.PlainLanguageSummary as string,
      FallbackContent: item.FallbackContent as string,
      ShortFormContent: item.ShortFormContent as string,
      RiskLevel: item.RiskLevel as ClauseRiskLevel,
      Negotiability: item.Negotiability as string,
      IsActive: item.IsActive as boolean,
      IsMandatory: item.IsMandatory as boolean,
      IsDefault: item.IsDefault as boolean,
      RegulatoryRequirement: item.RegulatoryRequirement as string,
      JurisdictionApplicability: item.JurisdictionApplicability as string,
      Variables: item.Variables as string,
      RelatedClauseIds: item.RelatedClauseIds as string,
      ConflictingClauseIds: item.ConflictingClauseIds as string,
      RequiresClauseIds: item.RequiresClauseIds as string,
      Version: item.Version as number || 1,
      EffectiveDate: item.EffectiveDate ? new Date(item.EffectiveDate as string) : undefined,
      RetiredDate: item.RetiredDate ? new Date(item.RetiredDate as string) : undefined,
      ReplacedByClauseId: item.ReplacedByClauseId as number,
      Author: item.Author as string,
      LegalReviewDate: item.LegalReviewDate ? new Date(item.LegalReviewDate as string) : undefined,
      LegalApprovedById: item.LegalApprovedById as number,
      UsageCount: item.UsageCount as number,
      Tags: item.Tags as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapClauseInstanceFromSP(item: Record<string, unknown>): IContractClauseInstance {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractId: item.ContractId as number,
      ClauseId: item.ClauseId as number,
      SectionNumber: item.SectionNumber as string,
      DisplayOrder: item.DisplayOrder as number,
      ClauseContent: item.ClauseContent as string,
      IsModified: item.IsModified as boolean,
      ModificationNotes: item.ModificationNotes as string,
      Status: item.Status as string,
      NegotiationNotes: item.NegotiationNotes as string,
      VariableValues: item.VariableValues as string,
      AddedById: item.AddedById as number,
      ModifiedById: item.ModifiedById as number,
      IsReviewed: item.IsReviewed as boolean,
      ReviewedById: item.ReviewedById as number,
      ReviewedDate: item.ReviewedDate ? new Date(item.ReviewedDate as string) : undefined,
      ReviewComments: item.ReviewComments as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapObligationFromSP(item: Record<string, unknown>): IContractObligation {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractId: item.ContractId as number,
      ObligationType: item.ObligationType as string,
      Description: item.Description as string,
      ClauseReference: item.ClauseReference as string,
      ResponsibleParty: item.ResponsibleParty as string,
      AssigneeId: item.AssigneeId as number,
      DueDate: new Date(item.DueDate as string),
      Frequency: item.Frequency as ObligationFrequency,
      RecurrencePattern: item.RecurrencePattern as string,
      NextOccurrence: item.NextOccurrence ? new Date(item.NextOccurrence as string) : undefined,
      EndDate: item.EndDate ? new Date(item.EndDate as string) : undefined,
      Status: item.Status as ObligationStatus,
      CompletedDate: item.CompletedDate ? new Date(item.CompletedDate as string) : undefined,
      CompletedById: item.CompletedById as number,
      CompletionNotes: item.CompletionNotes as string,
      ReminderDays: item.ReminderDays as number,
      RemindersSent: item.RemindersSent as number,
      LastReminderDate: item.LastReminderDate ? new Date(item.LastReminderDate as string) : undefined,
      Amount: item.Amount as number,
      Currency: item.Currency as Currency,
      EvidenceRequired: item.EvidenceRequired as boolean,
      EvidenceDocumentUrl: item.EvidenceDocumentUrl as string,
      HasPenalty: item.HasPenalty as boolean,
      PenaltyDescription: item.PenaltyDescription as string,
      PenaltyAmount: item.PenaltyAmount as number,
      Priority: item.Priority as ContractPriority,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapApprovalFromSP(item: Record<string, unknown>): IContractApproval {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractId: item.ContractId as number,
      ApprovalStage: item.ApprovalStage as number,
      ApprovalStageName: item.ApprovalStageName as string,
      ApproverId: item.ApproverId as number,
      DelegatedFromId: item.DelegatedFromId as number,
      DelegatedToId: item.DelegatedToId as number,
      Status: item.Status as ContractApprovalStatus,
      Action: item.Action as ContractApprovalAction,
      RequestedDate: new Date(item.RequestedDate as string),
      DueDate: item.DueDate ? new Date(item.DueDate as string) : undefined,
      ActionDate: item.ActionDate ? new Date(item.ActionDate as string) : undefined,
      RequestComments: item.RequestComments as string,
      ApprovalComments: item.ApprovalComments as string,
      ContractValue: item.ContractValue as number,
      ApprovalThreshold: item.ApprovalThreshold as number,
      RemindersSent: item.RemindersSent as number,
      LastReminderDate: item.LastReminderDate ? new Date(item.LastReminderDate as string) : undefined,
      FlowRunId: item.FlowRunId as string,
      FlowInstanceUrl: item.FlowInstanceUrl as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapApprovalRuleFromSP(item: Record<string, unknown>): IContractApprovalRule {
    return {
      Id: item.Id as number,
      Title: item.RuleName as string,
      ContractCategory: item.ContractCategory as ContractCategory,
      MinValue: item.MinValue as number,
      MaxValue: item.MaxValue as number,
      Department: item.Department as string,
      RiskLevel: item.RiskLevel as ContractRiskLevel,
      ApproverIds: item.ApproverIds as string,
      ApprovalOrder: item.ApprovalOrder as number,
      RequireAllApprovers: item.RequireAllApprovers as boolean,
      IsActive: item.IsActive as boolean,
      Priority: item.Priority as number,
      EscalationDays: item.EscalationDays as number,
      EscalateToId: item.EscalateToId as number,
      Notes: item.Notes as string
    };
  }

  private mapSignatureFromSP(item: Record<string, unknown>): IContractSignature {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractId: item.ContractId as number,
      ContractPartyId: item.ContractPartyId as number,
      Provider: item.Provider as string,
      ExternalEnvelopeId: item.ExternalEnvelopeId as string,
      ExternalSignerId: item.ExternalSignerId as string,
      Status: item.Status as SignatureStatus,
      RequestedDate: new Date(item.RequestedDate as string),
      SentDate: item.SentDate ? new Date(item.SentDate as string) : undefined,
      ViewedDate: item.ViewedDate ? new Date(item.ViewedDate as string) : undefined,
      SignedDate: item.SignedDate ? new Date(item.SignedDate as string) : undefined,
      DeclinedDate: item.DeclinedDate ? new Date(item.DeclinedDate as string) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate as string) : undefined,
      SignerName: item.SignerName as string,
      SignerEmail: item.SignerEmail as string,
      SignerTitle: item.SignerTitle as string,
      SignerCompany: item.SignerCompany as string,
      SignatureImageUrl: item.SignatureImageUrl as string,
      SignedDocumentUrl: item.SignedDocumentUrl as string,
      Certificate: item.Certificate as string,
      IPAddress: item.IPAddress as string,
      DeclineReason: item.DeclineReason as string,
      RemindersSent: item.RemindersSent as number,
      LastReminderDate: item.LastReminderDate ? new Date(item.LastReminderDate as string) : undefined,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapPartyFromSP(item: Record<string, unknown>): IContractParty {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractId: item.ContractId as number,
      PartyType: item.PartyType as string,
      PartyRole: item.PartyRole as string,
      UserId: item.UserId as number,
      VendorId: item.VendorId as number,
      ExternalName: item.ExternalName as string,
      ExternalTitle: item.ExternalTitle as string,
      ExternalCompany: item.ExternalCompany as string,
      ExternalEmail: item.ExternalEmail as string,
      ExternalPhone: item.ExternalPhone as string,
      IsSignatory: item.IsSignatory as boolean,
      SignatureOrder: item.SignatureOrder as number,
      SignatureStatus: item.SignatureStatus as SignatureStatus,
      SignedDate: item.SignedDate ? new Date(item.SignedDate as string) : undefined,
      SignatureId: item.SignatureId as string,
      Address: item.Address as string,
      City: item.City as string,
      State: item.State as string,
      Country: item.Country as string,
      PostalCode: item.PostalCode as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapTemplateFromSP(item: Record<string, unknown>): IContractTemplate {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      TemplateCode: item.TemplateCode as string,
      TemplateName: item.TemplateName as string,
      Category: item.Category as ContractCategory,
      Industry: item.Industry as string,
      Description: item.Description as string,
      UsageGuidance: item.UsageGuidance as string,
      TemplateDocumentUrl: item.TemplateDocumentUrl as string,
      DefaultClauses: item.DefaultClauses as string,
      MandatoryClauses: item.MandatoryClauses as string,
      DefaultDurationMonths: item.DefaultDurationMonths as number,
      DefaultRenewalType: item.DefaultRenewalType as ContractRenewalType,
      DefaultNotificationDays: item.DefaultNotificationDays as number,
      DefaultCurrency: item.DefaultCurrency as Currency,
      DefaultPaymentTerms: item.DefaultPaymentTerms as PaymentTerms,
      DefaultApproverIds: item.DefaultApproverIds as string,
      ApprovalThresholds: item.ApprovalThresholds as string,
      Variables: item.Variables as string,
      IsActive: item.IsActive as boolean,
      IsPublished: item.IsPublished as boolean,
      Version: item.Version as number,
      RequiresLegalReview: item.RequiresLegalReview as boolean,
      ComplianceChecklist: item.ComplianceChecklist as string,
      CreatedById: item.CreatedById as number,
      LastModifiedById: item.LastModifiedById as number,
      UsageCount: item.UsageCount as number,
      Tags: item.Tags as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapAuditLogFromSP(item: Record<string, unknown>): IContractAuditLog {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractId: item.ContractId as number,
      Action: item.Action as ContractAuditAction,
      ActionCategory: item.ActionCategory as string,
      ActionDescription: item.ActionDescription as string,
      PreviousValue: item.PreviousValue as string,
      NewValue: item.NewValue as string,
      ChangeDetails: item.ChangeDetails as string,
      RelatedEntityType: item.RelatedEntityType as string,
      RelatedEntityId: item.RelatedEntityId as number,
      ActionById: item.ActionById as number,
      ActionDate: new Date(item.ActionDate as string),
      IPAddress: item.IPAddress as string,
      UserAgent: item.UserAgent as string,
      SessionId: item.SessionId as string,
      IsSystemAction: item.IsSystemAction as boolean,
      Severity: item.Severity as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}

export default ContractManagementService;
