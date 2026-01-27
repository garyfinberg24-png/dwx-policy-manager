// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IJobOffer,
  OfferStatus,
  EmploymentType,
  ICandidate,
  CandidateStatus
} from '../models/ITalentManagement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

/**
 * Service for managing job offers, approvals, and acceptance tracking
 * Handles offer generation, approval workflows, and candidate communication
 */
export class OfferService {
  private sp: SPFI;
  private readonly JOB_OFFERS_LIST = 'Offers';
  private readonly CANDIDATES_LIST = 'Candidates';
  private readonly JOB_REQUISITIONS_LIST = 'Job Requisitions';
  private readonly CANDIDATE_ACTIVITIES_LIST = 'Candidate Activities';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== CRUD Operations ====================

  /**
   * Get job offers with optional filtering
   */
  public async getJobOffers(filter?: {
    candidateId?: number;
    jobRequisitionId?: number;
    status?: OfferStatus[];
    fromDate?: Date;
    toDate?: Date;
  }): Promise<IJobOffer[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.JOB_OFFERS_LIST).items
        .select(
          'Id',
          'CandidateId',
          'JobRequisitionId',
          'JobTitle',
          'Department',
          'Location',
          'EmploymentType',
          'StartDate',
          'BaseSalary',
          'Currency',
          'BonusAmount',
          'SigningBonus',
          'Equity',
          'Benefits',
          'IsRemote',
          'HybridSchedule',
          'WorkingHours',
          'ReportsToId',
          'ReportsTo/Title',
          'ReportsTo/EMail',
          'Status',
          'OfferDate',
          'ExpirationDate',
          'AcceptanceDate',
          'DeclineReason',
          'ApprovalRequired',
          'ApprovedById',
          'ApprovedBy/Title',
          'ApprovedBy/EMail',
          'ApprovalDate',
          'ApprovalNotes',
          'OfferLetterUrl',
          'ContractUrl',
          'IsSigned',
          'SignedDate',
          'OnboardingProcessId',
          'OnboardingStartDate',
          'Comments',
          'Attachments',
          'Created',
          'CreatedById',
          'Modified',
          'ModifiedById'
        )
        .expand('ReportsTo', 'ApprovedBy')
        .top(5000)
        .orderBy('Created', false);

      // Apply filters with SQL injection prevention
      if (filter) {
        const filters: string[] = [];

        if (filter.candidateId) {
          ValidationUtils.validateInteger(filter.candidateId, 'CandidateId', 1);
          filters.push(`CandidateId eq ${filter.candidateId}`);
        }

        if (filter.jobRequisitionId) {
          ValidationUtils.validateInteger(filter.jobRequisitionId, 'JobRequisitionId', 1);
          filters.push(`JobRequisitionId eq ${filter.jobRequisitionId}`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.fromDate) {
          ValidationUtils.validateDate(filter.fromDate, 'FromDate');
          filters.push(`OfferDate ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          ValidationUtils.validateDate(filter.toDate, 'ToDate');
          filters.push(`OfferDate le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query();
      return items.map(this.mapJobOfferFromSP);
    } catch (error) {
      logger.error('OfferService', 'Error getting job offers:', error);
      throw new Error(`Failed to retrieve job offers: ${error.message}`);
    }
  }

  /**
   * Get a single job offer by ID
   */
  public async getJobOfferById(id: number): Promise<IJobOffer> {
    try {
      ValidationUtils.validateInteger(id, 'Offer ID', 1);

      const item = await this.sp.web.lists.getByTitle(this.JOB_OFFERS_LIST).items
        .getById(id)
        .select(
          'Id',
          'CandidateId',
          'JobRequisitionId',
          'JobTitle',
          'Department',
          'Location',
          'EmploymentType',
          'StartDate',
          'BaseSalary',
          'Currency',
          'BonusAmount',
          'SigningBonus',
          'Equity',
          'Benefits',
          'IsRemote',
          'HybridSchedule',
          'WorkingHours',
          'ReportsToId',
          'ReportsTo/Title',
          'ReportsTo/EMail',
          'Status',
          'OfferDate',
          'ExpirationDate',
          'AcceptanceDate',
          'DeclineReason',
          'ApprovalRequired',
          'ApprovedById',
          'ApprovedBy/Title',
          'ApprovedBy/EMail',
          'ApprovalDate',
          'ApprovalNotes',
          'OfferLetterUrl',
          'ContractUrl',
          'IsSigned',
          'SignedDate',
          'OnboardingProcessId',
          'OnboardingStartDate',
          'Comments',
          'Attachments',
          'Created',
          'CreatedById',
          'Modified',
          'ModifiedById'
        )
        .expand('ReportsTo', 'ApprovedBy')();

      return this.mapJobOfferFromSP(item);
    } catch (error) {
      logger.error('OfferService', `Error getting job offer ${id}:`, error);
      throw new Error(`Failed to retrieve job offer: ${error.message}`);
    }
  }

  /**
   * Get offers for a specific candidate
   */
  public async getOffersForCandidate(candidateId: number): Promise<IJobOffer[]> {
    return this.getJobOffers({ candidateId });
  }

  /**
   * Create a new job offer
   */
  public async createJobOffer(offer: Partial<IJobOffer>): Promise<number> {
    try {
      // Validation
      if (!offer.CandidateId || !offer.JobRequisitionId || !offer.JobTitle || !offer.BaseSalary) {
        throw new Error('CandidateId, JobRequisitionId, JobTitle, and BaseSalary are required');
      }

      ValidationUtils.validateInteger(offer.CandidateId, 'CandidateId', 1);
      ValidationUtils.validateInteger(offer.JobRequisitionId, 'JobRequisitionId', 1);
      ValidationUtils.validateInteger(offer.BaseSalary, 'BaseSalary', 1);
      ValidationUtils.validateEnum(offer.EmploymentType, EmploymentType, 'EmploymentType');
      ValidationUtils.validateEnum(offer.Status, OfferStatus, 'Status');

      if (offer.BonusAmount) {
        ValidationUtils.validateInteger(offer.BonusAmount, 'BonusAmount', 0);
      }

      if (offer.SigningBonus) {
        ValidationUtils.validateInteger(offer.SigningBonus, 'SigningBonus', 0);
      }

      if (offer.StartDate) {
        ValidationUtils.validateDate(offer.StartDate, 'StartDate');
      }

      // Check if candidate already has an active offer
      const existingOffers = await this.getOffersForCandidate(offer.CandidateId);
      const activeOffer = existingOffers.find(o =>
        o.Status === OfferStatus.Sent ||
        o.Status === OfferStatus.Approved ||
        o.Status === OfferStatus.PendingApproval
      );

      if (activeOffer) {
        throw new Error(`Candidate already has an active offer (ID: ${activeOffer.Id}). Please withdraw or decline it first.`);
      }

      // Prepare item data
      const itemData: any = {
        CandidateId: offer.CandidateId,
        JobRequisitionId: offer.JobRequisitionId,
        JobTitle: ValidationUtils.sanitizeInput(offer.JobTitle),
        Department: offer.Department ? ValidationUtils.sanitizeInput(offer.Department) : null,
        Location: offer.Location ? ValidationUtils.sanitizeInput(offer.Location) : null,
        EmploymentType: offer.EmploymentType,
        StartDate: offer.StartDate || null,
        BaseSalary: offer.BaseSalary,
        Currency: offer.Currency || 'USD',
        BonusAmount: offer.BonusAmount || null,
        SigningBonus: offer.SigningBonus || null,
        Equity: offer.Equity ? ValidationUtils.sanitizeInput(offer.Equity) : null,
        Benefits: offer.Benefits || null,
        IsRemote: offer.IsRemote || false,
        HybridSchedule: offer.HybridSchedule ? ValidationUtils.sanitizeInput(offer.HybridSchedule) : null,
        WorkingHours: offer.WorkingHours ? ValidationUtils.sanitizeInput(offer.WorkingHours) : null,
        Status: offer.Status || OfferStatus.Draft,
        OfferDate: offer.OfferDate || null,
        ExpirationDate: offer.ExpirationDate || null,
        AcceptanceDate: null,
        DeclineReason: null,
        ApprovalRequired: offer.ApprovalRequired !== undefined ? offer.ApprovalRequired : true,
        ApprovalNotes: offer.ApprovalNotes ? ValidationUtils.sanitizeHtml(offer.ApprovalNotes) : null,
        OfferLetterUrl: offer.OfferLetterUrl ? ValidationUtils.sanitizeInput(offer.OfferLetterUrl) : null,
        ContractUrl: offer.ContractUrl ? ValidationUtils.sanitizeInput(offer.ContractUrl) : null,
        IsSigned: false,
        SignedDate: null,
        OnboardingProcessId: null,
        OnboardingStartDate: null,
        Notes: offer.Notes ? ValidationUtils.sanitizeHtml(offer.Notes) : null,
        Attachments: offer.Attachments || null
      };

      // Add lookup fields
      if (offer.ReportsToId) {
        ValidationUtils.validateInteger(offer.ReportsToId, 'ReportsToId', 1);
        itemData.ReportsToId = offer.ReportsToId;
      }
      if (offer.ApprovedById) {
        ValidationUtils.validateInteger(offer.ApprovedById, 'ApprovedById', 1);
        itemData.ApprovedById = offer.ApprovedById;
      }

      const result = await this.sp.web.lists.getByTitle(this.JOB_OFFERS_LIST).items.add(itemData);

      // Update candidate status to Offer
      await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items.getById(offer.CandidateId).update({
        Status: CandidateStatus.Offer,
        OfferId: result.data.Id
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer created',
        `Job offer created for ${offer.JobTitle} - ${offer.BaseSalary} ${offer.Currency || 'USD'}`
      );

      logger.debug('OfferService', `Job offer created successfully with ID: ${result.data.Id}`);
      return result.data.Id;
    } catch (error) {
      logger.error('OfferService', 'Error creating job offer:', error);
      throw new Error(`Failed to create job offer: ${error.message}`);
    }
  }

  /**
   * Update an existing job offer
   */
  public async updateJobOffer(id: number, updates: Partial<IJobOffer>): Promise<void> {
    try {
      ValidationUtils.validateInteger(id, 'Offer ID', 1);

      // Validate enums if provided
      if (updates.EmploymentType) {
        ValidationUtils.validateEnum(updates.EmploymentType, EmploymentType, 'EmploymentType');
      }
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, OfferStatus, 'Status');
      }

      if (updates.BaseSalary !== undefined) {
        ValidationUtils.validateInteger(updates.BaseSalary, 'BaseSalary', 1);
      }

      if (updates.BonusAmount !== undefined) {
        ValidationUtils.validateInteger(updates.BonusAmount, 'BonusAmount', 0);
      }

      if (updates.SigningBonus !== undefined) {
        ValidationUtils.validateInteger(updates.SigningBonus, 'SigningBonus', 0);
      }

      // Prepare update data
      const updateData: any = {};

      if (updates.CandidateId) {
        ValidationUtils.validateInteger(updates.CandidateId, 'CandidateId', 1);
        updateData.CandidateId = updates.CandidateId;
      }
      if (updates.JobRequisitionId) {
        ValidationUtils.validateInteger(updates.JobRequisitionId, 'JobRequisitionId', 1);
        updateData.JobRequisitionId = updates.JobRequisitionId;
      }
      if (updates.JobTitle) updateData.JobTitle = ValidationUtils.sanitizeInput(updates.JobTitle);
      if (updates.Department) updateData.Department = ValidationUtils.sanitizeInput(updates.Department);
      if (updates.Location) updateData.Location = ValidationUtils.sanitizeInput(updates.Location);
      if (updates.EmploymentType) updateData.EmploymentType = updates.EmploymentType;
      if (updates.StartDate) {
        ValidationUtils.validateDate(updates.StartDate, 'StartDate');
        updateData.StartDate = updates.StartDate;
      }
      if (updates.BaseSalary !== undefined) updateData.BaseSalary = updates.BaseSalary;
      if (updates.Currency) updateData.Currency = updates.Currency;
      if (updates.BonusAmount !== undefined) updateData.BonusAmount = updates.BonusAmount;
      if (updates.SigningBonus !== undefined) updateData.SigningBonus = updates.SigningBonus;
      if (updates.Equity) updateData.Equity = ValidationUtils.sanitizeInput(updates.Equity);
      if (updates.Benefits) updateData.Benefits = updates.Benefits;
      if (updates.IsRemote !== undefined) updateData.IsRemote = updates.IsRemote;
      if (updates.HybridSchedule) updateData.HybridSchedule = ValidationUtils.sanitizeInput(updates.HybridSchedule);
      if (updates.WorkingHours) updateData.WorkingHours = ValidationUtils.sanitizeInput(updates.WorkingHours);
      if (updates.Status) updateData.Status = updates.Status;
      if (updates.OfferDate) {
        ValidationUtils.validateDate(updates.OfferDate, 'OfferDate');
        updateData.OfferDate = updates.OfferDate;
      }
      if (updates.ExpirationDate) {
        ValidationUtils.validateDate(updates.ExpirationDate, 'ExpirationDate');
        updateData.ExpirationDate = updates.ExpirationDate;
      }
      if (updates.AcceptanceDate) {
        ValidationUtils.validateDate(updates.AcceptanceDate, 'AcceptanceDate');
        updateData.AcceptanceDate = updates.AcceptanceDate;
      }
      if (updates.DeclineReason) updateData.DeclineReason = ValidationUtils.sanitizeHtml(updates.DeclineReason);
      if (updates.ApprovalRequired !== undefined) updateData.ApprovalRequired = updates.ApprovalRequired;
      if (updates.ApprovalDate) {
        ValidationUtils.validateDate(updates.ApprovalDate, 'ApprovalDate');
        updateData.ApprovalDate = updates.ApprovalDate;
      }
      if (updates.ApprovalNotes) updateData.ApprovalNotes = ValidationUtils.sanitizeHtml(updates.ApprovalNotes);
      if (updates.OfferLetterUrl) updateData.OfferLetterUrl = ValidationUtils.sanitizeInput(updates.OfferLetterUrl);
      if (updates.ContractUrl) updateData.ContractUrl = ValidationUtils.sanitizeInput(updates.ContractUrl);
      if (updates.IsSigned !== undefined) updateData.IsSigned = updates.IsSigned;
      if (updates.SignedDate) {
        ValidationUtils.validateDate(updates.SignedDate, 'SignedDate');
        updateData.SignedDate = updates.SignedDate;
      }
      if (updates.OnboardingProcessId) {
        ValidationUtils.validateInteger(updates.OnboardingProcessId, 'OnboardingProcessId', 1);
        updateData.OnboardingProcessId = updates.OnboardingProcessId;
      }
      if (updates.OnboardingStartDate) {
        ValidationUtils.validateDate(updates.OnboardingStartDate, 'OnboardingStartDate');
        updateData.OnboardingStartDate = updates.OnboardingStartDate;
      }
      if (updates.Notes) updateData.Notes = ValidationUtils.sanitizeHtml(updates.Notes);
      if (updates.Attachments) updateData.Attachments = updates.Attachments;

      // Update lookup fields
      if (updates.ReportsToId) {
        ValidationUtils.validateInteger(updates.ReportsToId, 'ReportsToId', 1);
        updateData.ReportsToId = updates.ReportsToId;
      }
      if (updates.ApprovedById) {
        ValidationUtils.validateInteger(updates.ApprovedById, 'ApprovedById', 1);
        updateData.ApprovedById = updates.ApprovedById;
      }

      await this.sp.web.lists.getByTitle(this.JOB_OFFERS_LIST).items.getById(id).update(updateData);

      logger.debug('OfferService', `Job offer ${id} updated successfully`);
    } catch (error) {
      logger.error('OfferService', `Error updating job offer ${id}:`, error);
      throw new Error(`Failed to update job offer: ${error.message}`);
    }
  }

  /**
   * Delete a job offer
   */
  public async deleteJobOffer(id: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(id, 'Offer ID', 1);

      const offer = await this.getJobOfferById(id);

      if (offer.Status === OfferStatus.Accepted) {
        throw new Error('Cannot delete an accepted offer. Please withdraw it instead.');
      }

      await this.sp.web.lists.getByTitle(this.JOB_OFFERS_LIST).items.getById(id).delete();

      logger.debug('OfferService', `Job offer ${id} deleted successfully`);
    } catch (error) {
      logger.error('OfferService', `Error deleting job offer ${id}:`, error);
      throw new Error(`Failed to delete job offer: ${error.message}`);
    }
  }

  // ==================== Approval Workflow ====================

  /**
   * Submit offer for approval
   */
  public async submitForApproval(offerId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status !== OfferStatus.Draft) {
        throw new Error('Only draft offers can be submitted for approval');
      }

      if (!offer.ReportsToId) {
        throw new Error('Reporting manager must be assigned before submitting for approval');
      }

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.PendingApproval
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer submitted for approval',
        `Offer for ${offer.JobTitle} submitted for approval`
      );

      logger.debug('OfferService', `Offer ${offerId} submitted for approval`);
    } catch (error) {
      logger.error('OfferService', `Error submitting offer ${offerId} for approval:`, error);
      throw new Error(`Failed to submit for approval: ${error.message}`);
    }
  }

  /**
   * Approve a job offer
   */
  public async approveOffer(offerId: number, approverId: number, notes?: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);
      ValidationUtils.validateInteger(approverId, 'Approver ID', 1);

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status !== OfferStatus.PendingApproval) {
        throw new Error('Only offers pending approval can be approved');
      }

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Approved,
        ApprovedById: approverId,
        ApprovalDate: new Date(),
        ApprovalNotes: notes || 'Approved'
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer approved',
        `Offer for ${offer.JobTitle} approved${notes ? ': ' + notes : ''}`
      );

      logger.debug('OfferService', `Offer ${offerId} approved by user ${approverId}`);
    } catch (error) {
      logger.error('OfferService', `Error approving offer ${offerId}:`, error);
      throw new Error(`Failed to approve offer: ${error.message}`);
    }
  }

  /**
   * Reject a job offer
   */
  public async rejectOffer(offerId: number, reason: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      if (!reason) {
        throw new Error('Rejection reason is required');
      }

      const offer = await this.getJobOfferById(offerId);

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Draft,
        ApprovalNotes: ValidationUtils.sanitizeInput(reason)
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer rejected',
        `Offer rejected: ${reason}`
      );

      logger.debug('OfferService', `Offer ${offerId} rejected`);
    } catch (error) {
      logger.error('OfferService', `Error rejecting offer ${offerId}:`, error);
      throw new Error(`Failed to reject offer: ${error.message}`);
    }
  }

  // ==================== Offer Management ====================

  /**
   * Send offer to candidate
   */
  public async sendOffer(offerId: number, expirationDays?: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status !== OfferStatus.Approved) {
        throw new Error('Only approved offers can be sent to candidates');
      }

      const offerDate = new Date();
      const expirationDate = new Date();
      expirationDate.setDate(expirationDate.getDate() + (expirationDays || 7));

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Sent,
        OfferDate: offerDate,
        ExpirationDate: expirationDate
      });

      // Update candidate status
      await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items.getById(offer.CandidateId).update({
        Status: CandidateStatus.Offer,
        OfferDate: offerDate
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Email',
        'Offer sent',
        `Offer for ${offer.JobTitle} sent to candidate. Expires on ${expirationDate.toLocaleDateString()}`
      );

      logger.debug('OfferService', `Offer ${offerId} sent to candidate`);
    } catch (error) {
      logger.error('OfferService', `Error sending offer ${offerId}:`, error);
      throw new Error(`Failed to send offer: ${error.message}`);
    }
  }

  /**
   * Accept an offer (candidate acceptance)
   */
  public async acceptOffer(offerId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status !== OfferStatus.Sent) {
        throw new Error('Only sent offers can be accepted');
      }

      // Check if expired
      if (offer.ExpirationDate && new Date() > new Date(offer.ExpirationDate)) {
        throw new Error('This offer has expired');
      }

      const acceptanceDate = new Date();

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Accepted,
        AcceptanceDate: acceptanceDate
      });

      // Update candidate status to Hired
      await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items.getById(offer.CandidateId).update({
        Status: CandidateStatus.OfferAccepted,
        OfferDate: acceptanceDate
      });

      // Update job requisition filled positions
      const requisition = await this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items
        .getById(offer.JobRequisitionId)();

      const filledPositions = (requisition.FilledPositions || 0) + 1;

      await this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items
        .getById(offer.JobRequisitionId)
        .update({ FilledPositions: filledPositions });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer accepted',
        `Candidate accepted offer for ${offer.JobTitle}`
      );

      logger.debug('OfferService', `Offer ${offerId} accepted by candidate`);
    } catch (error) {
      logger.error('OfferService', `Error accepting offer ${offerId}:`, error);
      throw new Error(`Failed to accept offer: ${error.message}`);
    }
  }

  /**
   * Decline an offer (candidate decline)
   */
  public async declineOffer(offerId: number, reason?: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status !== OfferStatus.Sent) {
        throw new Error('Only sent offers can be declined');
      }

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Declined,
        DeclineReason: reason || 'Candidate declined offer'
      });

      // Update candidate status
      await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items.getById(offer.CandidateId).update({
        Status: CandidateStatus.OfferDeclined
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer declined',
        `Candidate declined offer for ${offer.JobTitle}${reason ? ': ' + reason : ''}`
      );

      logger.debug('OfferService', `Offer ${offerId} declined by candidate`);
    } catch (error) {
      logger.error('OfferService', `Error declining offer ${offerId}:`, error);
      throw new Error(`Failed to decline offer: ${error.message}`);
    }
  }

  /**
   * Withdraw an offer
   */
  public async withdrawOffer(offerId: number, reason: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      if (!reason) {
        throw new Error('Withdrawal reason is required');
      }

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status === OfferStatus.Accepted) {
        throw new Error('Cannot withdraw an accepted offer');
      }

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Withdrawn,
        DeclineReason: ValidationUtils.sanitizeInput(reason)
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer withdrawn',
        `Offer withdrawn: ${reason}`
      );

      logger.debug('OfferService', `Offer ${offerId} withdrawn`);
    } catch (error) {
      logger.error('OfferService', `Error withdrawing offer ${offerId}:`, error);
      throw new Error(`Failed to withdraw offer: ${error.message}`);
    }
  }

  /**
   * Mark offer as expired
   */
  public async markAsExpired(offerId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      const offer = await this.getJobOfferById(offerId);

      if (offer.Status !== OfferStatus.Sent) {
        throw new Error('Only sent offers can be marked as expired');
      }

      await this.updateJobOffer(offerId, {
        Status: OfferStatus.Expired
      });

      // Log activity
      await this.logCandidateActivity(
        offer.CandidateId,
        'Note',
        'Offer expired',
        `Offer for ${offer.JobTitle} has expired`
      );

      logger.debug('OfferService', `Offer ${offerId} marked as expired`);
    } catch (error) {
      logger.error('OfferService', `Error marking offer ${offerId} as expired:`, error);
      throw new Error(`Failed to mark as expired: ${error.message}`);
    }
  }

  /**
   * Check and expire offers automatically
   */
  public async checkAndExpireOffers(): Promise<number> {
    try {
      const sentOffers = await this.getJobOffers({
        status: [OfferStatus.Sent]
      });

      const now = new Date();
      let expiredCount = 0;

      for (const offer of sentOffers) {
        if (offer.ExpirationDate && new Date(offer.ExpirationDate) < now) {
          await this.markAsExpired(offer.Id);
          expiredCount++;
        }
      }

      if (expiredCount > 0) {
        logger.debug('OfferService', `Expired ${expiredCount} offers`);
      }

      return expiredCount;
    } catch (error) {
      logger.error('OfferService', 'Error checking and expiring offers:', error);
      throw new Error(`Failed to check and expire offers: ${error.message}`);
    }
  }

  // ==================== Analytics ====================

  /**
   * Get offer acceptance rate
   */
  public async getOfferAcceptanceRate(fromDate?: Date, toDate?: Date): Promise<number> {
    try {
      const filter: any = {};

      if (fromDate) filter.fromDate = fromDate;
      if (toDate) filter.toDate = toDate;

      const offers = await this.getJobOffers(filter);

      const sentOffers = offers.filter(o =>
        o.Status === OfferStatus.Sent ||
        o.Status === OfferStatus.Accepted ||
        o.Status === OfferStatus.Declined ||
        o.Status === OfferStatus.Expired
      );

      if (sentOffers.length === 0) return 0;

      const acceptedOffers = offers.filter(o => o.Status === OfferStatus.Accepted);

      return Math.round((acceptedOffers.length / sentOffers.length) * 100);
    } catch (error) {
      logger.error('OfferService', 'Error calculating offer acceptance rate:', error);
      throw new Error(`Failed to calculate acceptance rate: ${error.message}`);
    }
  }

  /**
   * Get average time to accept (days from sent to accepted)
   */
  public async getAverageTimeToAccept(): Promise<number> {
    try {
      const acceptedOffers = await this.getJobOffers({
        status: [OfferStatus.Accepted]
      });

      const offersWithDates = acceptedOffers.filter(o => o.OfferDate && o.AcceptanceDate);

      if (offersWithDates.length === 0) return 0;

      const totalDays = offersWithDates.reduce((sum, offer) => {
        const offerDate = new Date(offer.OfferDate).getTime();
        const acceptanceDate = new Date(offer.AcceptanceDate).getTime();
        const days = Math.floor((acceptanceDate - offerDate) / (1000 * 60 * 60 * 24));
        return sum + days;
      }, 0);

      return Math.round(totalDays / offersWithDates.length);
    } catch (error) {
      logger.error('OfferService', 'Error calculating average time to accept:', error);
      throw new Error(`Failed to calculate average time to accept: ${error.message}`);
    }
  }

  // ==================== Helper Methods ====================

  /**
   * Log candidate activity
   */
  private async logCandidateActivity(
    candidateId: number,
    activityType: 'Email' | 'Phone Call' | 'Meeting' | 'Interview' | 'Note' | 'Status Change' | 'Document Upload',
    subject?: string,
    description?: string
  ): Promise<number> {
    try {
      const activityData: any = {
        CandidateId: candidateId,
        ActivityType: activityType,
        ActivityDate: new Date(),
        Subject: subject ? ValidationUtils.sanitizeInput(subject) : null,
        Description: description ? ValidationUtils.sanitizeHtml(description) : null
      };

      const result = await this.sp.web.lists.getByTitle(this.CANDIDATE_ACTIVITIES_LIST).items.add(activityData);
      return result.data.Id;
    } catch (error) {
      logger.error('OfferService', 'Error logging candidate activity:', error);
      // Don't throw - activity logging should not block main operations
      return 0;
    }
  }

  /**
   * Map SharePoint item to IJobOffer
   */
  private mapJobOfferFromSP(item: any): IJobOffer {
    return {
      Id: item.Id,
      CandidateId: item.CandidateId,
      JobRequisitionId: item.JobRequisitionId,
      JobTitle: item.JobTitle,
      Department: item.Department,
      Location: item.Location,
      EmploymentType: item.EmploymentType,
      StartDate: item.StartDate ? new Date(item.StartDate) : undefined,
      BaseSalary: item.BaseSalary,
      Currency: item.Currency,
      BonusAmount: item.BonusAmount,
      SigningBonus: item.SigningBonus,
      Equity: item.Equity,
      Benefits: item.Benefits,
      IsRemote: item.IsRemote,
      HybridSchedule: item.HybridSchedule,
      WorkingHours: item.WorkingHours,
      ReportsToId: item.ReportsToId,
      ReportsTo: item.ReportsTo,
      Status: item.Status,
      OfferDate: item.OfferDate ? new Date(item.OfferDate) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined,
      AcceptanceDate: item.AcceptanceDate ? new Date(item.AcceptanceDate) : undefined,
      DeclineReason: item.DeclineReason,
      ApprovalRequired: item.ApprovalRequired,
      ApprovedById: item.ApprovedById,
      ApprovedBy: item.ApprovedBy,
      ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
      ApprovalNotes: item.ApprovalNotes,
      OfferLetterUrl: item.OfferLetterUrl,
      ContractUrl: item.ContractUrl,
      IsSigned: item.IsSigned,
      SignedDate: item.SignedDate ? new Date(item.SignedDate) : undefined,
      OnboardingProcessId: item.OnboardingProcessId,
      OnboardingStartDate: item.OnboardingStartDate ? new Date(item.OnboardingStartDate) : undefined,
      Notes: item.Notes,
      Attachments: item.Attachments,
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedById: item.CreatedById,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      ModifiedById: item.ModifiedById
    };
  }
}
