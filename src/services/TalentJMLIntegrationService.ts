// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IJobOffer,
  ICandidate,
  CandidateStatus
} from '../models/ITalentManagement';
import { OfferService } from './OfferService';
import { CandidateService } from './CandidateService';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

/**
 * Service for integrating Talent Management with JML Onboarding
 * Automates the handoff from hiring to employee onboarding
 */
export class TalentJMLIntegrationService {
  private sp: SPFI;
  private readonly JML_PROCESSES_LIST = 'JML Processes';
  private readonly M365_LICENSES_LIST = 'M365 Licenses';
  private readonly ASSET_ASSIGNMENTS_LIST = 'Asset Assignments';
  private readonly JOB_OFFERS_LIST = 'Offers';
  private readonly CANDIDATES_LIST = 'Candidates';
  private readonly CANDIDATE_ACTIVITIES_LIST = 'Candidate Activities';

  private offerService: OfferService;
  private candidateService: CandidateService;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.offerService = new OfferService(sp);
    this.candidateService = new CandidateService(sp);
  }

  // ==================== Main Integration Methods ====================

  /**
   * Create JML onboarding process when offer is accepted
   * This is the main integration point between Talent Management and JML
   */
  public async createOnboardingProcessFromOffer(offerId: number): Promise<number> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      // Get offer details
      const offer = await this.offerService.getJobOfferById(offerId);

      if (offer.Status !== 'Accepted') {
        throw new Error('Can only create onboarding process for accepted offers');
      }

      if (offer.OnboardingProcessId) {
        throw new Error(`Onboarding process already created (ID: ${offer.OnboardingProcessId})`);
      }

      // Get candidate details
      const candidate = await this.candidateService.getCandidateById(offer.CandidateId);

      // Create JML Process
      const processData: any = {
        Title: `Onboarding - ${candidate.FirstName} ${candidate.LastName}`,
        ProcessType: 'Joiner',
        EmployeeName: `${candidate.FirstName} ${candidate.LastName}`,
        EmployeeEmail: candidate.Email,
        Department: offer.Department,
        JobTitle: offer.JobTitle,
        Location: offer.Location,
        StartDate: offer.StartDate || new Date(),
        ManagerId: offer.ReportsToId,
        Status: 'Pending',
        Priority: 'High',
        Notes: `Created from accepted job offer (Offer ID: ${offerId})\n\nCandidate Information:\n- Phone: ${candidate.Phone || 'N/A'}\n- Current Company: ${candidate.CurrentCompany || 'N/A'}\n- Years of Experience: ${candidate.YearsOfExperience || 'N/A'}`,
        CreatedDate: new Date()
      };

      const result = await this.sp.web.lists.getByTitle(this.JML_PROCESSES_LIST).items.add(processData);
      const processId = result.data.Id;

      logger.debug('TalentJMLIntegrationService', `JML Process created with ID: ${processId}`);

      // Update offer with onboarding process ID
      await this.offerService.updateJobOffer(offerId, {
        OnboardingProcessId: processId,
        OnboardingStartDate: offer.StartDate || new Date()
      });

      // Update candidate status to Hired
      await this.candidateService.updateCandidate(candidate.Id, {
        Status: CandidateStatus.Hired
      });

      // Log activity
      await this.logCandidateActivity(
        candidate.Id,
        'Note',
        'Onboarding process created',
        `JML onboarding process created (Process ID: ${processId}). Start date: ${(offer.StartDate || new Date()).toLocaleDateString()}`
      );

      // Trigger additional onboarding tasks
      await this.setupOnboardingTasks(processId, offer, candidate);

      return processId;
    } catch (error) {
      logger.error('TalentJMLIntegrationService', `Error creating onboarding process from offer ${offerId}:`, error);
      throw new Error(`Failed to create onboarding process: ${error.message}`);
    }
  }

  /**
   * Setup onboarding tasks and preparations
   */
  private async setupOnboardingTasks(
    processId: number,
    offer: IJobOffer,
    candidate: ICandidate
  ): Promise<void> {
    try {
      // Request M365 license
      if (offer.JobTitle) {
        await this.requestM365License(processId, offer, candidate);
      }

      // Prepare asset assignment
      await this.prepareAssetAssignment(processId, offer, candidate);

      logger.debug('TalentJMLIntegrationService', `Onboarding tasks setup completed for process ${processId}`);
    } catch (error) {
      logger.error('TalentJMLIntegrationService', `Error setting up onboarding tasks for process ${processId}:`, error);
      // Don't throw - allow main process to succeed even if tasks fail
    }
  }

  /**
   * Request M365 license for new hire
   */
  private async requestM365License(
    processId: number,
    offer: IJobOffer,
    candidate: ICandidate
  ): Promise<number> {
    try {
      // Determine license type based on job title/department
      const licenseType = this.determineLicenseType(offer.JobTitle, offer.Department);

      const licenseData: any = {
        Title: `License Request - ${candidate.FirstName} ${candidate.LastName}`,
        UserEmail: candidate.Email,
        LicenseType: licenseType,
        RequestStatus: 'Pending',
        RequestDate: new Date(),
        RequiredByDate: offer.StartDate || new Date(),
        JMLProcessId: processId,
        Department: offer.Department,
        JobTitle: offer.JobTitle,
        Notes: `Auto-requested from talent management onboarding`,
        RequestedById: offer.CreatedById
      };

      const result = await this.sp.web.lists.getByTitle(this.M365_LICENSES_LIST).items.add(licenseData);

      logger.debug('TalentJMLIntegrationService', `M365 License request created with ID: ${result.data.Id}`);
      return result.data.Id;
    } catch (error) {
      logger.error('TalentJMLIntegrationService', 'Error requesting M365 license:', error);
      throw new Error(`Failed to request M365 license: ${error.message}`);
    }
  }

  /**
   * Prepare asset assignment for new hire
   */
  private async prepareAssetAssignment(
    processId: number,
    offer: IJobOffer,
    candidate: ICandidate
  ): Promise<void> {
    try {
      // Create asset assignment placeholder
      const assetData: any = {
        Title: `Asset Setup - ${candidate.FirstName} ${candidate.LastName}`,
        AssignedToEmail: candidate.Email,
        AssignmentDate: offer.StartDate || new Date(),
        JMLProcessId: processId,
        Department: offer.Department,
        Status: 'Pending',
        Notes: `Assets required for new hire starting ${(offer.StartDate || new Date()).toLocaleDateString()}`,
        IsActive: true
      };

      const result = await this.sp.web.lists.getByTitle(this.ASSET_ASSIGNMENTS_LIST).items.add(assetData);

      logger.debug('TalentJMLIntegrationService', `Asset assignment created with ID: ${result.data.Id}`);
    } catch (error) {
      logger.error('TalentJMLIntegrationService', 'Error creating asset assignment:', error);
      // Don't throw - this is a non-critical task
    }
  }

  /**
   * Determine appropriate M365 license type based on role
   */
  private determineLicenseType(jobTitle: string, department: string): string {
    const title = (jobTitle || '').toLowerCase();
    const dept = (department || '').toLowerCase();

    // Executive level
    if (title.includes('ceo') || title.includes('cto') || title.includes('cfo') || title.includes('president')) {
      return 'Microsoft 365 E5';
    }

    // Developer/Technical roles
    if (
      title.includes('developer') ||
      title.includes('engineer') ||
      title.includes('architect') ||
      dept.includes('engineering') ||
      dept.includes('development') ||
      dept.includes('it')
    ) {
      return 'Microsoft 365 E3';
    }

    // Sales/Marketing
    if (
      title.includes('sales') ||
      title.includes('marketing') ||
      dept.includes('sales') ||
      dept.includes('marketing')
    ) {
      return 'Microsoft 365 Business Premium';
    }

    // Management
    if (title.includes('manager') || title.includes('director') || title.includes('lead')) {
      return 'Microsoft 365 E3';
    }

    // Default for general staff
    return 'Microsoft 365 Business Standard';
  }

  /**
   * Sync candidate data to JML employee profile
   */
  public async syncCandidateToEmployee(candidateId: number, processId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(candidateId, 'Candidate ID', 1);
      ValidationUtils.validateInteger(processId, 'Process ID', 1);

      const candidate = await this.candidateService.getCandidateById(candidateId);

      // Update JML process with additional candidate information
      const updateData: any = {
        Notes: `${candidate.Notes || ''}\n\nCandidate Details:\n` +
          `- LinkedIn: ${candidate.LinkedInURL || 'N/A'}\n` +
          `- Skills: ${candidate.Skills || 'N/A'}\n` +
          `- Years of Experience: ${candidate.YearsOfExperience || 'N/A'}`
      };

      await this.sp.web.lists.getByTitle(this.JML_PROCESSES_LIST).items.getById(processId).update(updateData);

      logger.debug('TalentJMLIntegrationService', `Candidate data synced to JML process ${processId}`);
    } catch (error) {
      logger.error('TalentJMLIntegrationService', `Error syncing candidate ${candidateId} to JML:`, error);
      throw new Error(`Failed to sync candidate data: ${error.message}`);
    }
  }

  /**
   * Get onboarding status for a candidate
   */
  public async getOnboardingStatus(candidateId: number): Promise<{
    hasOnboardingProcess: boolean;
    processId?: number;
    processStatus?: string;
    startDate?: Date;
    completedTasks?: number;
    totalTasks?: number;
  }> {
    try {
      ValidationUtils.validateInteger(candidateId, 'Candidate ID', 1);

      const candidate = await this.candidateService.getCandidateById(candidateId);

      // Find offer for candidate
      const offersFilter = ValidationUtils.buildFilter('CandidateId', 'eq', candidateId.toString());
      const offers = await this.sp.web.lists.getByTitle(this.JOB_OFFERS_LIST).items
        .filter(offersFilter)
        .orderBy('Created', false)
        .top(1)();

      if (offers.length === 0 || !offers[0].OnboardingProcessId) {
        return {
          hasOnboardingProcess: false
        };
      }

      const processId = offers[0].OnboardingProcessId;

      // Get JML process details
      const process = await this.sp.web.lists.getByTitle(this.JML_PROCESSES_LIST).items
        .getById(processId)
        .select('Status', 'StartDate', 'CompletedTasks', 'TotalTasks')();

      return {
        hasOnboardingProcess: true,
        processId,
        processStatus: process.Status,
        startDate: process.StartDate ? new Date(process.StartDate) : undefined,
        completedTasks: process.CompletedTasks || 0,
        totalTasks: process.TotalTasks || 0
      };
    } catch (error) {
      logger.error('TalentJMLIntegrationService', `Error getting onboarding status for candidate ${candidateId}:`, error);
      throw new Error(`Failed to get onboarding status: ${error.message}`);
    }
  }

  /**
   * Bulk onboarding creation for multiple accepted offers
   */
  public async bulkCreateOnboardingProcesses(offerIds: number[]): Promise<{
    successful: number[];
    failed: { offerId: number; error: string }[];
  }> {
    const results = {
      successful: [] as number[],
      failed: [] as { offerId: number; error: string }[]
    };

    for (const offerId of offerIds) {
      try {
        const processId = await this.createOnboardingProcessFromOffer(offerId);
        results.successful.push(processId);
      } catch (error) {
        results.failed.push({
          offerId,
          error: error.message
        });
      }
    }

    return results;
  }

  /**
   * Validate offer readiness for onboarding
   */
  public async validateOfferForOnboarding(offerId: number): Promise<{
    isReady: boolean;
    issues: string[];
  }> {
    try {
      ValidationUtils.validateInteger(offerId, 'Offer ID', 1);

      const offer = await this.offerService.getJobOfferById(offerId);
      const issues: string[] = [];

      // Check offer status
      if (offer.Status !== 'Accepted') {
        issues.push('Offer must be accepted before creating onboarding process');
      }

      // Check required fields
      if (!offer.StartDate) {
        issues.push('Start date is required');
      }

      if (!offer.ReportsToId) {
        issues.push('Reporting manager must be assigned');
      }

      if (!offer.Department) {
        issues.push('Department is required');
      }

      if (!offer.JobTitle) {
        issues.push('Job title is required');
      }

      if (!offer.Location) {
        issues.push('Location is required');
      }

      // Check candidate information
      const candidate = await this.candidateService.getCandidateById(offer.CandidateId);

      if (!candidate.Email) {
        issues.push('Candidate email is required');
      }

      if (!candidate.Phone) {
        issues.push('Candidate phone number is required (recommended)');
      }

      return {
        isReady: issues.length === 0,
        issues
      };
    } catch (error) {
      logger.error('TalentJMLIntegrationService', `Error validating offer ${offerId} for onboarding:`, error);
      throw new Error(`Failed to validate offer: ${error.message}`);
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

      const result = await this.sp.web.lists
        .getByTitle(this.CANDIDATE_ACTIVITIES_LIST)
        .items.add(activityData);
      return result.data.Id;
    } catch (error) {
      logger.error('TalentJMLIntegrationService', 'Error logging candidate activity:', error);
      // Don't throw - activity logging should not block main operations
      return 0;
    }
  }

  /**
   * Get integration metrics
   */
  public async getIntegrationMetrics(): Promise<{
    totalOnboardingProcesses: number;
    pendingOnboarding: number;
    completedOnboarding: number;
    averageDaysToOnboard: number;
  }> {
    try {
      // Get all JML processes created from talent management
      const processes = await this.sp.web.lists.getByTitle(this.JML_PROCESSES_LIST).items
        .filter("ProcessType eq 'Joiner' and Title ne null")
        .select('Id', 'Status', 'StartDate', 'Created', 'Modified')
        .top(5000)();

      const totalOnboardingProcesses = processes.length;
      const pendingOnboarding = processes.filter(p => p.Status === 'Pending' || p.Status === 'In Progress').length;
      const completedOnboarding = processes.filter(p => p.Status === 'Completed').length;

      // Calculate average days to onboard
      const completedProcesses = processes.filter(p => p.Status === 'Completed' && p.StartDate && p.Modified);
      let totalDays = 0;

      completedProcesses.forEach(process => {
        const startDate = new Date(process.StartDate);
        const completedDate = new Date(process.Modified);
        const days = Math.floor((completedDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));
        totalDays += days;
      });

      const averageDaysToOnboard = completedProcesses.length > 0
        ? Math.round(totalDays / completedProcesses.length)
        : 0;

      return {
        totalOnboardingProcesses,
        pendingOnboarding,
        completedOnboarding,
        averageDaysToOnboard
      };
    } catch (error) {
      logger.error('TalentJMLIntegrationService', 'Error getting integration metrics:', error);
      throw new Error(`Failed to get integration metrics: ${error.message}`);
    }
  }
}
