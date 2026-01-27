// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IJobRequisition,
  JobRequisitionStatus,
  EmploymentType,
  Priority,
  IRecruitmentMetrics,
  ApplicationSource,
  CandidateStatus
} from '../models/ITalentManagement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

/**
 * Service for managing job requisitions and recruitment operations
 * Handles position tracking, approval workflows, and recruitment metrics
 */
export class RecruitmentService {
  private sp: SPFI;
  private readonly JOB_REQUISITIONS_LIST = 'Job Requisitions';
  private readonly CANDIDATES_LIST = 'Candidates';
  private readonly INTERVIEWS_LIST = 'Interviews';
  private readonly JOB_OFFERS_LIST = 'Offers';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== CRUD Operations ====================

  /**
   * Get job requisitions with optional filtering
   */
  public async getJobRequisitions(filter?: {
    status?: JobRequisitionStatus[];
    department?: string;
    hiringManagerId?: number;
    recruiterId?: number;
    priority?: Priority[];
    employmentType?: EmploymentType[];
    isPublished?: boolean;
    searchTerm?: string;
  }): Promise<IJobRequisition[]> {
    try {
      // Query only fields that exist in SharePoint list - Person fields may not exist
      let query = this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items
        .select(
          'Id',
          'Title',
          'JobCode',
          'Department',
          'Location',
          'EmploymentType',
          'Status',
          'NumberOfOpenings',
          'FilledPositions',
          'Priority',
          'IsPublished',
          'PostedDate',
          'Created',
          'Modified'
        )
        .top(5000)
        .orderBy('Created', false);

      // Apply filters with SQL injection prevention
      if (filter) {
        const filters: string[] = [];

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.department) {
          filters.push(ValidationUtils.buildFilter('Department', 'eq', filter.department));
        }

        if (filter.hiringManagerId) {
          ValidationUtils.validateInteger(filter.hiringManagerId, 'HiringManagerId', 1);
          filters.push(`HiringManagerId eq ${filter.hiringManagerId}`);
        }

        if (filter.recruiterId) {
          ValidationUtils.validateInteger(filter.recruiterId, 'RecruiterId', 1);
          filters.push(`RecruiterId eq ${filter.recruiterId}`);
        }

        if (filter.priority && filter.priority.length > 0) {
          const priorityFilters = filter.priority.map(p =>
            ValidationUtils.buildFilter('Priority', 'eq', p)
          );
          filters.push(`(${priorityFilters.join(' or ')})`);
        }

        if (filter.employmentType && filter.employmentType.length > 0) {
          const typeFilters = filter.employmentType.map(t =>
            ValidationUtils.buildFilter('EmploymentType', 'eq', t)
          );
          filters.push(`(${typeFilters.join(' or ')})`);
        }

        if (filter.isPublished !== undefined) {
          filters.push(`IsPublished eq ${filter.isPublished ? '1' : '0'}`);
        }

        if (filter.searchTerm) {
          const sanitized = ValidationUtils.sanitizeInput(filter.searchTerm);
          filters.push(`(substringof('${sanitized}', Title) or substringof('${sanitized}', JobDescription))`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query();
      return items.map(this.mapJobRequisitionFromSP);
    } catch (error) {
      logger.error('RecruitmentService', 'Error getting job requisitions:', error);
      throw new Error(`Failed to retrieve job requisitions: ${error.message}`);
    }
  }

  /**
   * Get a single job requisition by ID
   */
  public async getJobRequisitionById(id: number): Promise<IJobRequisition> {
    try {
      ValidationUtils.validateInteger(id, 'Requisition ID', 1);

      // Query only fields that exist in SharePoint list - Person fields may not exist
      const item = await this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items
        .getById(id)
        .select(
          'Id',
          'Title',
          'JobCode',
          'Department',
          'Location',
          'EmploymentType',
          'Status',
          'NumberOfOpenings',
          'FilledPositions',
          'Priority',
          'IsPublished',
          'PostedDate',
          'Created',
          'Modified'
        )();

      return this.mapJobRequisitionFromSP(item);
    } catch (error) {
      logger.error('RecruitmentService', `Error getting job requisition ${id}:`, error);
      throw new Error(`Failed to retrieve job requisition: ${error.message}`);
    }
  }

  /**
   * Create a new job requisition
   */
  public async createJobRequisition(requisition: Partial<IJobRequisition>): Promise<number> {
    try {
      // Validation
      if (!requisition.Title || !requisition.Department || !requisition.Location) {
        throw new Error('Title, Department, and Location are required');
      }

      ValidationUtils.validateEnum(requisition.EmploymentType, EmploymentType, 'EmploymentType');
      ValidationUtils.validateEnum(requisition.Status, JobRequisitionStatus, 'Status');
      ValidationUtils.validateEnum(requisition.Priority, Priority, 'Priority');
      ValidationUtils.validateInteger(requisition.NumberOfOpenings, 'NumberOfOpenings', 1);

      if (requisition.SalaryRangeMin) {
        ValidationUtils.validateInteger(requisition.SalaryRangeMin, 'SalaryRangeMin', 0);
      }

      if (requisition.SalaryRangeMax) {
        ValidationUtils.validateInteger(requisition.SalaryRangeMax, 'SalaryRangeMax', 0);
        if (requisition.SalaryRangeMin && requisition.SalaryRangeMax < requisition.SalaryRangeMin) {
          throw new Error('SalaryRangeMax cannot be less than SalaryRangeMin');
        }
      }

      if (requisition.HiringManagerId) {
        ValidationUtils.validateInteger(requisition.HiringManagerId, 'HiringManagerId', 1);
      }

      if (requisition.RecruiterId) {
        ValidationUtils.validateInteger(requisition.RecruiterId, 'RecruiterId', 1);
      }

      // Prepare item data
      const itemData: any = {
        Title: ValidationUtils.sanitizeInput(requisition.Title),
        JobCode: requisition.JobCode ? ValidationUtils.sanitizeInput(requisition.JobCode) : null,
        Department: ValidationUtils.sanitizeInput(requisition.Department),
        Location: ValidationUtils.sanitizeInput(requisition.Location),
        EmploymentType: requisition.EmploymentType,
        Status: requisition.Status || JobRequisitionStatus.Draft,
        NumberOfOpenings: requisition.NumberOfOpenings,
        FilledPositions: requisition.FilledPositions || 0,
        Priority: requisition.Priority || Priority.Medium,
        JobDescription: requisition.JobDescription ? ValidationUtils.sanitizeHtml(requisition.JobDescription) : '',
        Responsibilities: requisition.Responsibilities ? ValidationUtils.sanitizeHtml(requisition.Responsibilities) : null,
        Requirements: requisition.Requirements ? ValidationUtils.sanitizeHtml(requisition.Requirements) : null,
        PreferredQualifications: requisition.PreferredQualifications ? ValidationUtils.sanitizeHtml(requisition.PreferredQualifications) : null,
        Skills: requisition.Skills ? ValidationUtils.sanitizeInput(requisition.Skills) : null,
        Certifications: requisition.Certifications ? ValidationUtils.sanitizeInput(requisition.Certifications) : null,
        SalaryRangeMin: requisition.SalaryRangeMin || null,
        SalaryRangeMax: requisition.SalaryRangeMax || null,
        Currency: requisition.Currency || 'USD',
        BonusEligible: requisition.BonusEligible || false,
        BenefitsDescription: requisition.BenefitsDescription ? ValidationUtils.sanitizeHtml(requisition.BenefitsDescription) : null,
        ApprovalRequired: requisition.ApprovalRequired || false,
        BudgetCode: requisition.BudgetCode ? ValidationUtils.sanitizeInput(requisition.BudgetCode) : null,
        EstimatedTotalCost: requisition.EstimatedTotalCost || null,
        IsPublished: requisition.IsPublished || false,
        ExternalJobBoards: requisition.ExternalJobBoards || null,
        Notes: requisition.Notes ? ValidationUtils.sanitizeHtml(requisition.Notes) : null,
        Attachments: requisition.Attachments || null
      };

      // Add lookup fields
      if (requisition.HiringManagerId) {
        itemData.HiringManagerId = requisition.HiringManagerId;
      }
      if (requisition.RecruiterId) {
        itemData.RecruiterId = requisition.RecruiterId;
      }

      // Add date fields
      if (requisition.TargetStartDate) {
        ValidationUtils.validateDate(requisition.TargetStartDate, 'TargetStartDate');
        itemData.TargetStartDate = requisition.TargetStartDate;
      }
      if (requisition.TargetFillDate) {
        ValidationUtils.validateDate(requisition.TargetFillDate, 'TargetFillDate');
        itemData.TargetFillDate = requisition.TargetFillDate;
      }

      const result = await this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items.add(itemData);

      logger.debug('RecruitmentService', `Job requisition created successfully with ID: ${result.data.Id}`);
      return result.data.Id;
    } catch (error) {
      logger.error('RecruitmentService', 'Error creating job requisition:', error);
      throw new Error(`Failed to create job requisition: ${error.message}`);
    }
  }

  /**
   * Update an existing job requisition
   */
  public async updateJobRequisition(id: number, updates: Partial<IJobRequisition>): Promise<void> {
    try {
      ValidationUtils.validateInteger(id, 'Requisition ID', 1);

      // Validate enums if provided
      if (updates.EmploymentType) {
        ValidationUtils.validateEnum(updates.EmploymentType, EmploymentType, 'EmploymentType');
      }
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, JobRequisitionStatus, 'Status');
      }
      if (updates.Priority) {
        ValidationUtils.validateEnum(updates.Priority, Priority, 'Priority');
      }

      if (updates.NumberOfOpenings !== undefined) {
        ValidationUtils.validateInteger(updates.NumberOfOpenings, 'NumberOfOpenings', 1);
      }

      if (updates.SalaryRangeMin !== undefined) {
        ValidationUtils.validateInteger(updates.SalaryRangeMin, 'SalaryRangeMin', 0);
      }

      if (updates.SalaryRangeMax !== undefined) {
        ValidationUtils.validateInteger(updates.SalaryRangeMax, 'SalaryRangeMax', 0);
      }

      // Prepare update data
      const updateData: any = {};

      if (updates.Title) updateData.Title = ValidationUtils.sanitizeInput(updates.Title);
      if (updates.JobCode) updateData.JobCode = ValidationUtils.sanitizeInput(updates.JobCode);
      if (updates.Department) updateData.Department = ValidationUtils.sanitizeInput(updates.Department);
      if (updates.Location) updateData.Location = ValidationUtils.sanitizeInput(updates.Location);
      if (updates.EmploymentType) updateData.EmploymentType = updates.EmploymentType;
      if (updates.Status) updateData.Status = updates.Status;
      if (updates.Priority) updateData.Priority = updates.Priority;
      if (updates.NumberOfOpenings !== undefined) updateData.NumberOfOpenings = updates.NumberOfOpenings;
      if (updates.FilledPositions !== undefined) updateData.FilledPositions = updates.FilledPositions;
      if (updates.JobDescription) updateData.JobDescription = ValidationUtils.sanitizeHtml(updates.JobDescription);
      if (updates.Responsibilities) updateData.Responsibilities = ValidationUtils.sanitizeHtml(updates.Responsibilities);
      if (updates.Requirements) updateData.Requirements = ValidationUtils.sanitizeHtml(updates.Requirements);
      if (updates.PreferredQualifications) updateData.PreferredQualifications = ValidationUtils.sanitizeHtml(updates.PreferredQualifications);
      if (updates.Skills) updateData.Skills = ValidationUtils.sanitizeInput(updates.Skills);
      if (updates.Certifications) updateData.Certifications = ValidationUtils.sanitizeInput(updates.Certifications);
      if (updates.SalaryRangeMin !== undefined) updateData.SalaryRangeMin = updates.SalaryRangeMin;
      if (updates.SalaryRangeMax !== undefined) updateData.SalaryRangeMax = updates.SalaryRangeMax;
      if (updates.Currency) updateData.Currency = updates.Currency;
      if (updates.BonusEligible !== undefined) updateData.BonusEligible = updates.BonusEligible;
      if (updates.BenefitsDescription) updateData.BenefitsDescription = ValidationUtils.sanitizeHtml(updates.BenefitsDescription);
      if (updates.ApprovalRequired !== undefined) updateData.ApprovalRequired = updates.ApprovalRequired;
      if (updates.ApprovalNotes) updateData.ApprovalNotes = ValidationUtils.sanitizeHtml(updates.ApprovalNotes);
      if (updates.BudgetCode) updateData.BudgetCode = ValidationUtils.sanitizeInput(updates.BudgetCode);
      if (updates.EstimatedTotalCost !== undefined) updateData.EstimatedTotalCost = updates.EstimatedTotalCost;
      if (updates.IsPublished !== undefined) updateData.IsPublished = updates.IsPublished;
      if (updates.ExternalJobBoards) updateData.ExternalJobBoards = updates.ExternalJobBoards;
      if (updates.Notes) updateData.Notes = ValidationUtils.sanitizeHtml(updates.Notes);
      if (updates.Attachments) updateData.Attachments = updates.Attachments;

      // Update lookup fields
      if (updates.HiringManagerId) {
        ValidationUtils.validateInteger(updates.HiringManagerId, 'HiringManagerId', 1);
        updateData.HiringManagerId = updates.HiringManagerId;
      }
      if (updates.RecruiterId) {
        ValidationUtils.validateInteger(updates.RecruiterId, 'RecruiterId', 1);
        updateData.RecruiterId = updates.RecruiterId;
      }
      if (updates.ApprovedById) {
        ValidationUtils.validateInteger(updates.ApprovedById, 'ApprovedById', 1);
        updateData.ApprovedById = updates.ApprovedById;
      }

      // Update date fields
      if (updates.TargetStartDate) {
        ValidationUtils.validateDate(updates.TargetStartDate, 'TargetStartDate');
        updateData.TargetStartDate = updates.TargetStartDate;
      }
      if (updates.TargetFillDate) {
        ValidationUtils.validateDate(updates.TargetFillDate, 'TargetFillDate');
        updateData.TargetFillDate = updates.TargetFillDate;
      }
      if (updates.ApprovalDate) {
        ValidationUtils.validateDate(updates.ApprovalDate, 'ApprovalDate');
        updateData.ApprovalDate = updates.ApprovalDate;
      }
      if (updates.PostedDate) {
        ValidationUtils.validateDate(updates.PostedDate, 'PostedDate');
        updateData.PostedDate = updates.PostedDate;
      }
      if (updates.PostingExpiration) {
        ValidationUtils.validateDate(updates.PostingExpiration, 'PostingExpiration');
        updateData.PostingExpiration = updates.PostingExpiration;
      }

      await this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items.getById(id).update(updateData);

      logger.debug('RecruitmentService', `Job requisition ${id} updated successfully`);
    } catch (error) {
      logger.error('RecruitmentService', `Error updating job requisition ${id}:`, error);
      throw new Error(`Failed to update job requisition: ${error.message}`);
    }
  }

  /**
   * Delete a job requisition
   */
  public async deleteJobRequisition(id: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(id, 'Requisition ID', 1);

      // Check if there are candidates associated
      const candidatesFilter = ValidationUtils.buildFilter('JobRequisitionId', 'eq', id.toString());
      const candidates = await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items
        .filter(candidatesFilter)
        .top(1)();

      if (candidates.length > 0) {
        throw new Error('Cannot delete job requisition with associated candidates. Please archive or close instead.');
      }

      await this.sp.web.lists.getByTitle(this.JOB_REQUISITIONS_LIST).items.getById(id).delete();

      logger.debug('RecruitmentService', `Job requisition ${id} deleted successfully`);
    } catch (error) {
      logger.error('RecruitmentService', `Error deleting job requisition ${id}:`, error);
      throw new Error(`Failed to delete job requisition: ${error.message}`);
    }
  }

  // ==================== Approval Workflow ====================

  /**
   * Submit requisition for approval
   */
  public async submitForApproval(requisitionId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(requisitionId, 'Requisition ID', 1);

      const requisition = await this.getJobRequisitionById(requisitionId);

      if (requisition.Status !== JobRequisitionStatus.Draft) {
        throw new Error('Only draft requisitions can be submitted for approval');
      }

      if (!requisition.HiringManagerId) {
        throw new Error('Hiring Manager must be assigned before submitting for approval');
      }

      await this.updateJobRequisition(requisitionId, {
        Status: JobRequisitionStatus.PendingApproval
      });

      logger.debug('RecruitmentService', `Job requisition ${requisitionId} submitted for approval`);
    } catch (error) {
      logger.error('RecruitmentService', `Error submitting requisition ${requisitionId} for approval:`, error);
      throw new Error(`Failed to submit for approval: ${error.message}`);
    }
  }

  /**
   * Approve a job requisition
   */
  public async approveRequisition(requisitionId: number, approverId: number, notes?: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(requisitionId, 'Requisition ID', 1);
      ValidationUtils.validateInteger(approverId, 'Approver ID', 1);

      const requisition = await this.getJobRequisitionById(requisitionId);

      if (requisition.Status !== JobRequisitionStatus.PendingApproval) {
        throw new Error('Only requisitions pending approval can be approved');
      }

      await this.updateJobRequisition(requisitionId, {
        Status: JobRequisitionStatus.Approved,
        ApprovedById: approverId,
        ApprovalDate: new Date(),
        ApprovalNotes: notes || 'Approved'
      });

      logger.debug('RecruitmentService', `Job requisition ${requisitionId} approved by user ${approverId}`);
    } catch (error) {
      logger.error('RecruitmentService', `Error approving requisition ${requisitionId}:`, error);
      throw new Error(`Failed to approve requisition: ${error.message}`);
    }
  }

  /**
   * Reject a job requisition
   */
  public async rejectRequisition(requisitionId: number, reason: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(requisitionId, 'Requisition ID', 1);

      if (!reason) {
        throw new Error('Rejection reason is required');
      }

      await this.updateJobRequisition(requisitionId, {
        Status: JobRequisitionStatus.Draft,
        ApprovalNotes: ValidationUtils.sanitizeInput(reason)
      });

      logger.debug('RecruitmentService', `Job requisition ${requisitionId} rejected`);
    } catch (error) {
      logger.error('RecruitmentService', `Error rejecting requisition ${requisitionId}:`, error);
      throw new Error(`Failed to reject requisition: ${error.message}`);
    }
  }

  /**
   * Publish a job requisition
   */
  public async publishRequisition(requisitionId: number, externalJobBoards?: string[]): Promise<void> {
    try {
      ValidationUtils.validateInteger(requisitionId, 'Requisition ID', 1);

      const requisition = await this.getJobRequisitionById(requisitionId);

      if (requisition.Status !== JobRequisitionStatus.Approved) {
        throw new Error('Only approved requisitions can be published');
      }

      await this.updateJobRequisition(requisitionId, {
        Status: JobRequisitionStatus.Open,
        IsPublished: true,
        PostedDate: new Date(),
        ExternalJobBoards: externalJobBoards ? JSON.stringify(externalJobBoards) : null
      });

      logger.debug('RecruitmentService', `Job requisition ${requisitionId} published`);
    } catch (error) {
      logger.error('RecruitmentService', `Error publishing requisition ${requisitionId}:`, error);
      throw new Error(`Failed to publish requisition: ${error.message}`);
    }
  }

  /**
   * Close a filled requisition
   */
  public async closeRequisition(requisitionId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(requisitionId, 'Requisition ID', 1);

      const requisition = await this.getJobRequisitionById(requisitionId);

      if (requisition.FilledPositions >= requisition.NumberOfOpenings) {
        await this.updateJobRequisition(requisitionId, {
          Status: JobRequisitionStatus.Filled
        });
        logger.debug('RecruitmentService', `Job requisition ${requisitionId} marked as filled`);
      } else {
        throw new Error(`Cannot close requisition: Only ${requisition.FilledPositions} of ${requisition.NumberOfOpenings} positions filled`);
      }
    } catch (error) {
      logger.error('RecruitmentService', `Error closing requisition ${requisitionId}:`, error);
      throw new Error(`Failed to close requisition: ${error.message}`);
    }
  }

  // ==================== Analytics & Metrics ====================

  /**
   * Helper method to safely get list items, returning empty array if list doesn't exist or has field errors
   */
  private async safeGetListItems<T>(listName: string, selectFields: string[]): Promise<T[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(listName).items
        .select(...selectFields)
        .top(5000)();
      return items as T[];
    } catch (error: any) {
      // Check if list doesn't exist (404 error or list not found) OR field doesn't exist (400 Bad Request)
      if (error.status === 404 || error.status === 400 ||
          error.message?.includes('does not exist') ||
          error.message?.includes('404') ||
          error.message?.includes('400')) {
        logger.warn('RecruitmentService', `List '${listName}' not found or has field errors - returning empty array`);
        return [];
      }
      throw error;
    }
  }

  /**
   * Get comprehensive recruitment metrics
   */
  public async getRecruitmentMetrics(): Promise<IRecruitmentMetrics> {
    try {
      // Get all requisitions (handle missing list gracefully)
      const requisitions = await this.safeGetListItems<any>(
        this.JOB_REQUISITIONS_LIST,
        ['Id', 'Status', 'Created', 'NumberOfOpenings', 'FilledPositions', 'PostedDate', 'Modified']
      );

      // Get all candidates (handle missing list gracefully) - use minimal fields
      const candidates = await this.safeGetListItems<any>(
        this.CANDIDATES_LIST,
        ['Id', 'Title', 'Created', 'Modified']
      );

      // Get all interviews (handle missing list gracefully) - use minimal fields
      const interviews = await this.safeGetListItems<any>(
        this.INTERVIEWS_LIST,
        ['Id', 'Title', 'Created', 'Modified']
      );

      // Get all offers (handle missing list gracefully) - use minimal fields
      const offers = await this.safeGetListItems<any>(
        this.JOB_OFFERS_LIST,
        ['Id', 'Title', 'Created', 'Modified']
      );

      // Calculate metrics - using basic counts since detailed fields may not exist
      const totalRequisitions = requisitions.length;
      const openRequisitions = requisitions.filter(r => r.Status === JobRequisitionStatus.Open).length;
      const filledRequisitions = requisitions.filter(r => r.Status === JobRequisitionStatus.Filled).length;

      const totalCandidates = candidates.length;
      const newCandidates = 0; // Cannot determine without Status field

      // Candidates by status - empty since we don't have Status field
      const candidatesByStatus: { [key in CandidateStatus]?: number } = {};

      // Candidates by source - empty since we don't have Source field
      const candidatesBySource: { [key in ApplicationSource]?: number } = {};

      // Best performing source - undefined since we don't have Source data
      const bestPerformingSource: ApplicationSource | undefined = undefined;

      // Interview metrics - basic counts only
      const totalInterviews = interviews.length;
      const upcomingInterviews = 0; // Cannot determine without ScheduledDate field
      const completedInterviews = 0; // Cannot determine without Status field
      const interviewCompletionRate = 0;

      // Offer metrics - basic counts only
      const totalOffers = offers.length;
      const pendingOffers = 0; // Cannot determine without Status field
      const acceptedOffers = 0; // Cannot determine without Status field
      const declinedOffers = 0; // Cannot determine without Status field
      const offerAcceptanceRate = 0;

      // Time to fill (average days from posted to filled)
      let totalDaysToFill = 0;
      let filledCount = 0;
      requisitions.filter(r => r.Status === JobRequisitionStatus.Filled && r.PostedDate).forEach(req => {
        const postedDate = new Date(req.PostedDate);
        const modifiedDate = new Date(req.Modified);
        const days = Math.floor((modifiedDate.getTime() - postedDate.getTime()) / (1000 * 60 * 60 * 24));
        totalDaysToFill += days;
        filledCount++;
      });
      const avgTimeToFill = filledCount > 0 ? Math.round(totalDaysToFill / filledCount) : 0;

      // Time to hire - cannot calculate without detailed fields
      const avgTimeToHire = 0;

      // Average ratings - cannot calculate without rating fields
      const avgCandidateRating = 0;
      const avgInterviewScore = 0;

      // Recent activity (last 7 days) - based on Created date only
      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

      const recentActivity = {
        newApplications: candidates.filter(c => new Date(c.Created) >= sevenDaysAgo).length,
        scheduledInterviews: interviews.filter(i => new Date(i.Created) >= sevenDaysAgo).length,
        completedInterviews: 0, // Cannot determine without Status field
        offersExtended: offers.filter(o => new Date(o.Created) >= sevenDaysAgo).length
      };

      const metrics: IRecruitmentMetrics = {
        totalCandidates,
        newCandidates,
        candidatesByStatus,
        totalRequisitions,
        openRequisitions,
        filledRequisitions,
        avgTimeToFill,
        avgTimeToHire,
        totalInterviews,
        upcomingInterviews,
        completedInterviews,
        interviewCompletionRate,
        totalOffers,
        pendingOffers,
        acceptedOffers,
        declinedOffers,
        offerAcceptanceRate,
        candidatesBySource,
        bestPerformingSource,
        avgCandidateRating,
        avgInterviewScore,
        totalRecruitmentCost: 0, // Would need additional cost tracking
        costPerHire: 0,
        referralBonusesPaid: 0,
        recentActivity
      };

      return metrics;
    } catch (error) {
      logger.error('RecruitmentService', 'Error getting recruitment metrics:', error);
      throw new Error(`Failed to retrieve recruitment metrics: ${error.message}`);
    }
  }

  /**
   * Get time to fill analytics for specific requisition
   */
  public async getTimeToFillForRequisition(requisitionId: number): Promise<number> {
    try {
      ValidationUtils.validateInteger(requisitionId, 'Requisition ID', 1);

      const requisition = await this.getJobRequisitionById(requisitionId);

      if (!requisition.PostedDate) {
        return 0;
      }

      const publishedDate = new Date(requisition.PostedDate);
      const currentDate = requisition.Status === JobRequisitionStatus.Filled
        ? new Date(requisition.Modified)
        : new Date();

      const days = Math.floor((currentDate.getTime() - publishedDate.getTime()) / (1000 * 60 * 60 * 24));

      return days;
    } catch (error) {
      logger.error('RecruitmentService', `Error calculating time to fill for requisition ${requisitionId}:`, error);
      throw new Error(`Failed to calculate time to fill: ${error.message}`);
    }
  }

  /**
   * Get open positions by department
   */
  public async getOpenPositionsByDepartment(): Promise<{ department: string; count: number }[]> {
    try {
      const openRequisitions = await this.getJobRequisitions({
        status: [JobRequisitionStatus.Open, JobRequisitionStatus.Approved]
      });

      const departmentMap = new Map<string, number>();

      openRequisitions.forEach(req => {
        const current = departmentMap.get(req.Department) || 0;
        departmentMap.set(req.Department, current + req.NumberOfOpenings - (req.FilledPositions || 0));
      });

      return Array.from(departmentMap.entries()).map(([department, count]) => ({
        department,
        count
      })).sort((a, b) => b.count - a.count);
    } catch (error) {
      logger.error('RecruitmentService', 'Error getting open positions by department:', error);
      throw new Error(`Failed to get open positions by department: ${error.message}`);
    }
  }

  // ==================== Helper Methods ====================

  /**
   * Map SharePoint item to IJobRequisition
   */
  private mapJobRequisitionFromSP(item: any): IJobRequisition {
    return {
      Id: item.Id,
      Title: item.Title,
      JobCode: item.JobCode,
      Department: item.Department,
      Location: item.Location,
      EmploymentType: item.EmploymentType,
      Status: item.Status,
      HiringManagerId: item.HiringManagerId,
      HiringManager: item.HiringManager,
      RecruiterId: item.RecruiterId,
      Recruiter: item.Recruiter,
      NumberOfOpenings: item.NumberOfOpenings,
      FilledPositions: item.FilledPositions || 0,
      Priority: item.Priority,
      TargetStartDate: item.TargetStartDate ? new Date(item.TargetStartDate) : undefined,
      TargetFillDate: item.TargetFillDate ? new Date(item.TargetFillDate) : undefined,
      JobDescription: item.JobDescription,
      Responsibilities: item.Responsibilities,
      Requirements: item.Requirements,
      PreferredQualifications: item.PreferredQualifications,
      Skills: item.Skills,
      Certifications: item.Certifications,
      SalaryRangeMin: item.SalaryRangeMin,
      SalaryRangeMax: item.SalaryRangeMax,
      Currency: item.Currency,
      BonusEligible: item.BonusEligible,
      BenefitsDescription: item.BenefitsDescription,
      ApprovalRequired: item.ApprovalRequired,
      ApprovedById: item.ApprovedById,
      ApprovedBy: item.ApprovedBy,
      ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
      ApprovalNotes: item.ApprovalNotes,
      BudgetCode: item.BudgetCode,
      EstimatedTotalCost: item.EstimatedTotalCost,
      IsPublished: item.IsPublished,
      PostedDate: item.PostedDate ? new Date(item.PostedDate) : undefined,
      PostingExpiration: item.PostingExpiration ? new Date(item.PostingExpiration) : undefined,
      ExternalJobBoards: item.ExternalJobBoards,
      Notes: item.Notes,
      Attachments: item.Attachments,
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedById: item.CreatedById,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      ModifiedById: item.ModifiedById
    };
  }
}
