// @ts-nocheck
// Candidate Service
// Comprehensive candidate tracking and management

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  ICandidate,
  ICandidateActivity,
  ICandidateFilterCriteria,
  ICandidateScoreCard,
  CandidateStatus,
  ApplicationSource
} from '../models/ITalentManagement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class CandidateService {
  private sp: SPFI;
  private readonly CANDIDATES_LIST = 'Candidates';
  private readonly CANDIDATE_ACTIVITIES_LIST = 'Candidate Activities';
  private readonly INTERVIEWS_LIST = 'Interviews';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Candidate CRUD Operations ====================

  public async getCandidates(filter?: ICandidateFilterCriteria): Promise<ICandidate[]> {
    try {
      // Note: Field names must match SharePoint list internal names from provisioning script
      // CurrentJobTitle (not CurrentPosition), LinkedInUrl (not LinkedInURL), ApplicationSource (not Source)
      let query = this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items
        .select(
          'Id', 'FirstName', 'LastName', 'Email', 'Phone', 'LinkedInUrl',
          'CurrentJobTitle', 'CurrentCompany', 'YearsOfExperience',
          'Status', 'ApplicationDate', 'ApplicationSource',
          'JobRequisitionId', 'ResumeUrl', 'CoverLetterUrl',
          'OverallRating',
          'Skills', 'Notes',
          'Created', 'Modified'
        )
        .orderBy('ApplicationDate', false);

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.jobRequisitionId !== undefined) {
          const validReqId = ValidationUtils.validateInteger(filter.jobRequisitionId, 'jobRequisitionId', 1);
          filters.push(ValidationUtils.buildFilter('JobRequisitionId', 'eq', validReqId));
        }

        if (filter.source && filter.source.length > 0) {
          const sourceFilters = filter.source.map(s =>
            ValidationUtils.buildFilter('Source', 'eq', s)
          );
          filters.push(`(${sourceFilters.join(' or ')})`);
        }

        if (filter.minRating !== undefined) {
          const validRating = ValidationUtils.validateInteger(filter.minRating, 'minRating', 1, 5);
          filters.push(ValidationUtils.buildFilter('OverallRating', 'ge', validRating));
        }

        if (filter.location) {
          const validLocation = ValidationUtils.sanitizeForOData(filter.location);
          filters.push(`(substringof('${validLocation}', City) or substringof('${validLocation}', State))`);
        }

        if (filter.minExperience !== undefined) {
          const validMinExp = ValidationUtils.validateInteger(filter.minExperience, 'minExperience', 0);
          filters.push(ValidationUtils.buildFilter('YearsOfExperience', 'ge', validMinExp));
        }

        if (filter.maxExperience !== undefined) {
          const validMaxExp = ValidationUtils.validateInteger(filter.maxExperience, 'maxExperience', 0);
          filters.push(ValidationUtils.buildFilter('YearsOfExperience', 'le', validMaxExp));
        }

        if (filter.searchTerm) {
          const validTerm = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${validTerm}', FirstName) or substringof('${validTerm}', LastName) or substringof('${validTerm}', Email) or substringof('${validTerm}', CurrentCompany) or substringof('${validTerm}', Skills))`);
        }

        if (filter.fromDate) {
          ValidationUtils.validateDate(filter.fromDate, 'fromDate');
          filters.push(ValidationUtils.buildFilter('ApplicationDate', 'ge', filter.fromDate));
        }

        if (filter.toDate) {
          ValidationUtils.validateDate(filter.toDate, 'toDate');
          filters.push(ValidationUtils.buildFilter('ApplicationDate', 'le', filter.toDate));
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.top(5000)();
      return items.map(this.mapCandidateFromSP);
    } catch (error) {
      logger.error('CandidateService', 'Error getting candidates:', error);
      throw error;
    }
  }

  public async getCandidateById(id: number): Promise<ICandidate> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Note: Field names must match SharePoint list internal names from provisioning script
      const item = await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items
        .getById(validId)
        .select(
          'Id', 'FirstName', 'LastName', 'Email', 'Phone', 'LinkedInUrl',
          'CurrentCompany', 'CurrentJobTitle', 'YearsOfExperience',
          'Status', 'ApplicationDate', 'ApplicationSource',
          'JobRequisitionId', 'ResumeUrl', 'CoverLetterUrl',
          'OverallRating',
          'ReferredById', 'ReferredBy/Title', 'ReferredBy/EMail',
          'Skills',
          'RejectionReason', 'RejectionDate', 'IsEligibleForRehire',
          'LastContactDate', 'NextFollowUpDate',
          'Gender', 'Ethnicity', 'Veteran', 'Disability',
          'Notes', 'Tags', 'Created', 'Modified'
        )
        .expand('ReferredBy')();

      return this.mapCandidateFromSP(item);
    } catch (error) {
      logger.error('CandidateService', 'Error getting candidate by ID:', error);
      throw error;
    }
  }

  public async createCandidate(candidate: Partial<ICandidate>): Promise<number> {
    try {
      // Validate required fields
      if (!candidate.FirstName || !candidate.LastName || !candidate.Email) {
        throw new Error('FirstName, LastName, and Email are required');
      }

      // Validate email
      ValidationUtils.validateEmail(candidate.Email, 'Email');

      // Validate status
      if (candidate.Status) {
        ValidationUtils.validateEnum(candidate.Status, CandidateStatus, 'Status');
      }

      if (candidate.ApplicationSource) {
        ValidationUtils.validateEnum(candidate.ApplicationSource, ApplicationSource, 'ApplicationSource');
      }

      // Check for duplicate email
      const existingFilter = ValidationUtils.buildFilter('Email', 'eq', candidate.Email);
      const existing = await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items
        .filter(existingFilter)
        .top(1)();

      if (existing.length > 0) {
        throw new Error(`Candidate with email ${candidate.Email} already exists`);
      }

      const itemData: any = {
        FirstName: ValidationUtils.sanitizeHtml(candidate.FirstName),
        LastName: ValidationUtils.sanitizeHtml(candidate.LastName),
        Email: ValidationUtils.sanitizeHtml(candidate.Email),
        Status: candidate.Status || CandidateStatus.New,
        ApplicationDate: candidate.ApplicationDate || new Date(),
        ApplicationSource: candidate.ApplicationSource || ApplicationSource.DirectApply
      };

      // Optional fields
      if (candidate.Phone) itemData.Phone = ValidationUtils.sanitizeHtml(candidate.Phone);
      if (candidate.LinkedInURL) itemData.LinkedInURL = ValidationUtils.sanitizeHtml(candidate.LinkedInURL);
      if (candidate.CurrentCompany) itemData.CurrentCompany = ValidationUtils.sanitizeHtml(candidate.CurrentCompany);
      if (candidate.CurrentPosition) itemData.CurrentPosition = ValidationUtils.sanitizeHtml(candidate.CurrentPosition);
      if (candidate.YearsOfExperience !== undefined) itemData.YearsOfExperience = ValidationUtils.validateInteger(candidate.YearsOfExperience, 'YearsOfExperience', 0);
      if (candidate.City) itemData.City = ValidationUtils.sanitizeHtml(candidate.City);
      if (candidate.State) itemData.State = ValidationUtils.sanitizeHtml(candidate.State);
      if (candidate.Country) itemData.Country = ValidationUtils.sanitizeHtml(candidate.Country);
      if (candidate.IsWillingToRelocate !== undefined) itemData.IsWillingToRelocate = candidate.IsWillingToRelocate;
      if (candidate.JobRequisitionId) itemData.JobRequisitionId = ValidationUtils.validateInteger(candidate.JobRequisitionId, 'JobRequisitionId', 1);
      if (candidate.ReferredById) itemData.ReferredById = ValidationUtils.validateInteger(candidate.ReferredById, 'ReferredById', 1);
      if (candidate.ResumeURL) itemData.ResumeURL = ValidationUtils.sanitizeHtml(candidate.ResumeURL);
      if (candidate.CoverLetterURL) itemData.CoverLetterURL = ValidationUtils.sanitizeHtml(candidate.CoverLetterURL);
      if (candidate.Skills) itemData.Skills = ValidationUtils.sanitizeHtml(candidate.Skills);
      if (candidate.Certifications) itemData.Certifications = ValidationUtils.sanitizeHtml(candidate.Certifications);
      if (candidate.Education) itemData.Education = ValidationUtils.sanitizeHtml(candidate.Education);
      if (candidate.ExpectedSalary !== undefined) itemData.ExpectedSalary = ValidationUtils.validateInteger(candidate.ExpectedSalary, 'ExpectedSalary', 0);
      if (candidate.NoticePeriod !== undefined) itemData.NoticePeriod = ValidationUtils.validateInteger(candidate.NoticePeriod, 'NoticePeriod', 0);
      if (candidate.EarliestStartDate) itemData.EarliestStartDate = ValidationUtils.validateDate(candidate.EarliestStartDate, 'EarliestStartDate');
      if (candidate.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(candidate.Notes);
      if (candidate.Tags) itemData.Tags = ValidationUtils.sanitizeHtml(candidate.Tags);

      const result = await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items.add(itemData);

      // Log activity
      await this.logActivity(result.data.Id, 'Note', 'Candidate profile created', `Application received from ${candidate.ApplicationSource}`, undefined);

      return result.data.Id;
    } catch (error) {
      logger.error('CandidateService', 'Error creating candidate:', error);
      throw error;
    }
  }

  public async updateCandidate(id: number, updates: Partial<ICandidate>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: any = {};

      if (updates.FirstName) itemData.FirstName = ValidationUtils.sanitizeHtml(updates.FirstName);
      if (updates.LastName) itemData.LastName = ValidationUtils.sanitizeHtml(updates.LastName);
      if (updates.Email) {
        ValidationUtils.validateEmail(updates.Email, 'Email');
        itemData.Email = ValidationUtils.sanitizeHtml(updates.Email);
      }
      if (updates.Phone) itemData.Phone = ValidationUtils.sanitizeHtml(updates.Phone);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, CandidateStatus, 'Status');
        itemData.Status = updates.Status;

        // Log status change
        await this.logActivity(validId, 'Status Change', `Status changed to ${updates.Status}`, updates.Notes, undefined);
      }
      if (updates.JobRequisitionId) itemData.JobRequisitionId = ValidationUtils.validateInteger(updates.JobRequisitionId, 'JobRequisitionId', 1);
      if (updates.OverallRating !== undefined) itemData.OverallRating = ValidationUtils.validateInteger(updates.OverallRating, 'OverallRating', 1, 5);
      if (updates.TechnicalSkillsRating !== undefined) itemData.TechnicalSkillsRating = ValidationUtils.validateInteger(updates.TechnicalSkillsRating, 'TechnicalSkillsRating', 1, 5);
      if (updates.CommunicationSkillsRating !== undefined) itemData.CommunicationSkillsRating = ValidationUtils.validateInteger(updates.CommunicationSkillsRating, 'CommunicationSkillsRating', 1, 5);
      if (updates.CulturalFitRating !== undefined) itemData.CulturalFitRating = ValidationUtils.validateInteger(updates.CulturalFitRating, 'CulturalFitRating', 1, 5);
      if (updates.BackgroundCheckStatus) itemData.BackgroundCheckStatus = updates.BackgroundCheckStatus;
      if (updates.BackgroundCheckDate) itemData.BackgroundCheckDate = ValidationUtils.validateDate(updates.BackgroundCheckDate, 'BackgroundCheckDate');
      if (updates.OfferId !== undefined) itemData.OfferId = updates.OfferId;
      if (updates.OfferAmount !== undefined) itemData.OfferAmount = updates.OfferAmount;
      if (updates.OfferDate) itemData.OfferDate = ValidationUtils.validateDate(updates.OfferDate, 'OfferDate');
      if (updates.RejectionReason) {
        itemData.RejectionReason = ValidationUtils.sanitizeHtml(updates.RejectionReason);
        itemData.RejectionDate = new Date();
      }
      if (updates.IsEligibleForRehire !== undefined) itemData.IsEligibleForRehire = updates.IsEligibleForRehire;
      if (updates.LastContactDate) itemData.LastContactDate = ValidationUtils.validateDate(updates.LastContactDate, 'LastContactDate');
      if (updates.NextFollowUpDate) itemData.NextFollowUpDate = ValidationUtils.validateDate(updates.NextFollowUpDate, 'NextFollowUpDate');
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.Tags !== undefined) itemData.Tags = updates.Tags ? ValidationUtils.sanitizeHtml(updates.Tags) : null;

      await this.sp.web.lists.getByTitle(this.CANDIDATES_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('CandidateService', 'Error updating candidate:', error);
      throw error;
    }
  }

  public async deleteCandidate(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Soft delete by marking as withdrawn
      await this.updateCandidate(validId, {
        Status: CandidateStatus.Withdrawn,
        Notes: 'Candidate record deleted'
      });
    } catch (error) {
      logger.error('CandidateService', 'Error deleting candidate:', error);
      throw error;
    }
  }

  // ==================== Candidate Activities ====================

  public async logActivity(
    candidateId: number,
    activityType: 'Email' | 'Phone Call' | 'Meeting' | 'Interview' | 'Note' | 'Status Change' | 'Document Upload',
    subject?: string,
    description?: string,
    performedById?: number
  ): Promise<number> {
    try {
      const validCandidateId = ValidationUtils.validateInteger(candidateId, 'candidateId', 1);

      const activityData: any = {
        CandidateId: validCandidateId,
        ActivityType: activityType,
        ActivityDate: new Date().toISOString(),
        Subject: subject ? ValidationUtils.sanitizeHtml(subject) : null,
        Description: description ? ValidationUtils.sanitizeHtml(description) : null,
        PerformedById: performedById
      };

      const result = await this.sp.web.lists.getByTitle(this.CANDIDATE_ACTIVITIES_LIST).items.add(activityData);
      return result.data.Id;
    } catch (error) {
      logger.error('CandidateService', 'Error logging activity:', error);
      throw error;
    }
  }

  public async getActivities(candidateId: number): Promise<ICandidateActivity[]> {
    try {
      const validCandidateId = ValidationUtils.validateInteger(candidateId, 'candidateId', 1);

      const filter = ValidationUtils.buildFilter('CandidateId', 'eq', validCandidateId);

      const items = await this.sp.web.lists.getByTitle(this.CANDIDATE_ACTIVITIES_LIST).items
        .filter(filter)
        .select(
          'Id', 'CandidateId', 'ActivityType', 'ActivityDate', 'Subject', 'Description',
          'PerformedById', 'PerformedBy/Title', 'RelatedInterviewId', 'Attachments', 'Created'
        )
        .expand('PerformedBy')
        .orderBy('ActivityDate', false)
        .top(500)();

      return items.map(this.mapActivityFromSP);
    } catch (error) {
      logger.error('CandidateService', 'Error getting activities:', error);
      throw error;
    }
  }

  // ==================== Candidate Scorecard ====================

  public async getCandidateScoreCard(candidateId: number): Promise<ICandidateScoreCard> {
    try {
      const validCandidateId = ValidationUtils.validateInteger(candidateId, 'candidateId', 1);

      // Get candidate
      const candidate = await this.getCandidateById(validCandidateId);

      // Get all interviews for this candidate
      const interviewFilter = ValidationUtils.buildFilter('CandidateId', 'eq', validCandidateId);
      const interviews = await this.sp.web.lists.getByTitle(this.INTERVIEWS_LIST).items
        .filter(interviewFilter)
        .select(
          'Id', 'InterviewType', 'OverallScore', 'InterviewerId', 'Interviewer/Title', 'ScheduledDate', 'Result'
        )
        .expand('Interviewer')();

      const interviewScores = interviews
        .filter(i => i.OverallScore)
        .map(i => ({
          interviewId: i.Id,
          interviewType: i.InterviewType,
          score: i.OverallScore,
          interviewer: i.Interviewer?.Title || 'Unknown',
          date: new Date(i.ScheduledDate)
        }));

      const technicalInterviews = interviews.filter(i => i.InterviewType === 'Technical Interview' && i.OverallScore);
      const behavioralInterviews = interviews.filter(i => i.InterviewType === 'Behavioral Interview' && i.OverallScore);

      const technicalAverage = technicalInterviews.length > 0
        ? technicalInterviews.reduce((sum, i) => sum + i.OverallScore, 0) / technicalInterviews.length
        : 0;

      const behavioralAverage = behavioralInterviews.length > 0
        ? behavioralInterviews.reduce((sum, i) => sum + i.OverallScore, 0) / behavioralInterviews.length
        : 0;

      const overallScore = candidate.OverallRating || 0;

      let recommendation: 'Strong Hire' | 'Hire' | 'Maybe' | 'No Hire' = 'Maybe';
      if (overallScore >= 4.5) recommendation = 'Strong Hire';
      else if (overallScore >= 3.5) recommendation = 'Hire';
      else if (overallScore >= 2.5) recommendation = 'Maybe';
      else recommendation = 'No Hire';

      return {
        candidateId: validCandidateId,
        candidateName: `${candidate.FirstName} ${candidate.LastName}`,
        overallScore: overallScore,
        interviewScores: interviewScores,
        technicalAverage: technicalAverage,
        behavioralAverage: behavioralAverage,
        recommendation: recommendation
      };
    } catch (error) {
      logger.error('CandidateService', 'Error getting candidate scorecard:', error);
      throw error;
    }
  }

  // ==================== Mapping Functions ====================

  private mapCandidateFromSP(item: any): ICandidate {
    // Map SharePoint field names to ICandidate interface
    // SP uses: CurrentJobTitle, LinkedInUrl, ApplicationSource, ResumeUrl, CoverLetterUrl
    return {
      Id: item.Id,
      FirstName: item.FirstName,
      LastName: item.LastName,
      Email: item.Email,
      Phone: item.Phone,
      LinkedInURL: item.LinkedInUrl, // SP: LinkedInUrl
      CurrentCompany: item.CurrentCompany,
      CurrentPosition: item.CurrentJobTitle, // SP: CurrentJobTitle
      YearsOfExperience: item.YearsOfExperience,
      HighestDegree: item.HighestDegree,
      Status: item.Status as CandidateStatus,
      ApplicationDate: item.ApplicationDate ? new Date(item.ApplicationDate) : new Date(),
      ApplicationSource: item.ApplicationSource as ApplicationSource, // SP: ApplicationSource
      ReferredById: item.ReferredById,
      ReferredBy: item.ReferredBy,
      ReferralBonus: item.ReferralBonus,
      JobRequisitionId: item.JobRequisitionId,
      ResumeURL: item.ResumeUrl, // SP: ResumeUrl
      CoverLetterURL: item.CoverLetterUrl, // SP: CoverLetterUrl
      AdditionalDocuments: item.AdditionalDocuments,
      OverallRating: item.OverallRating,
      TechnicalSkillsRating: item.TechnicalSkillsRating,
      CommunicationSkillsRating: item.CommunicationSkillsRating,
      CulturalFitRating: item.CulturalFitRating,
      Skills: item.Skills,
      Certifications: item.Certifications,
      Education: item.Education,
      Languages: item.Languages,
      ExpectedSalary: item.ExpectedSalary,
      CurrentSalary: item.CurrentSalary,
      NoticePeriod: item.NoticePeriod,
      EarliestStartDate: item.EarliestStartDate ? new Date(item.EarliestStartDate) : undefined,
      BackgroundCheckStatus: item.BackgroundCheckStatus,
      BackgroundCheckDate: item.BackgroundCheckDate ? new Date(item.BackgroundCheckDate) : undefined,
      BackgroundCheckNotes: item.BackgroundCheckNotes,
      OfferId: item.OfferId,
      OfferAmount: item.OfferAmount,
      OfferDate: item.OfferDate ? new Date(item.OfferDate) : undefined,
      RejectionReason: item.RejectionReason,
      RejectionDate: item.RejectionDate ? new Date(item.RejectionDate) : undefined,
      IsEligibleForRehire: item.IsEligibleForRehire,
      LastContactDate: item.LastContactDate ? new Date(item.LastContactDate) : undefined,
      NextFollowUpDate: item.NextFollowUpDate ? new Date(item.NextFollowUpDate) : undefined,
      Gender: item.Gender,
      Ethnicity: item.Ethnicity,
      Veteran: item.Veteran,
      Disability: item.Disability,
      Notes: item.Notes,
      Tags: item.Tags,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapActivityFromSP(item: any): ICandidateActivity {
    return {
      Id: item.Id,
      CandidateId: item.CandidateId,
      ActivityType: item.ActivityType,
      ActivityDate: new Date(item.ActivityDate),
      Subject: item.Subject,
      Description: item.Description,
      PerformedById: item.PerformedById,
      PerformedBy: item.PerformedBy,
      RelatedInterviewId: item.RelatedInterviewId,
      Attachments: item.Attachments,
      Created: item.Created ? new Date(item.Created) : undefined
    };
  }
}
