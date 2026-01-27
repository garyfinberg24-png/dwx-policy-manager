// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IInterview,
  InterviewType,
  InterviewStatus,
  InterviewResult,
  IInterviewFilterCriteria,
  IScheduleConflict,
  ICandidateActivity
} from '../models/ITalentManagement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

/**
 * Service for managing interviews, scheduling, and feedback collection
 * Handles interview coordination, conflict detection, and evaluation tracking
 */
export class InterviewService {
  private sp: SPFI;
  private readonly INTERVIEWS_LIST = 'Interviews';
  private readonly CANDIDATES_LIST = 'Candidates';
  private readonly CANDIDATE_ACTIVITIES_LIST = 'Candidate Activities';
  private readonly INTERVIEW_FEEDBACK_TEMPLATES_LIST = 'Interview Feedback Templates';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== CRUD Operations ====================

  /**
   * Get interviews with optional filtering
   */
  public async getInterviews(filter?: IInterviewFilterCriteria): Promise<IInterview[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.INTERVIEWS_LIST).items
        .select(
          'Id',
          'CandidateId',
          'JobRequisitionId',
          'InterviewType',
          'ScheduledDate',
          'Duration',
          'Status',
          'Location',
          'MeetingLink',
          'RoomNumber',
          'IsVirtual',
          'InterviewerId',
          'Interviewer/Title',
          'Interviewer/EMail',
          'PanelMembers',
          'CoordinatorId',
          'Coordinator/Title',
          'Result',
          'OverallScore',
          'TechnicalScore',
          'CommunicationScore',
          'ProblemSolvingScore',
          'CulturalFitScore',
          'Feedback',
          'Strengths',
          'Weaknesses',
          'Recommendation',
          'HiringRecommendation',
          'QuestionsAsked',
          'CandidateQuestions',
          'TechnicalAssessmentUrl',
          'PresentationUrl',
          'NextSteps',
          'FollowUpRequired',
          'FollowUpDate',
          'CandidateNotified',
          'ReminderSent',
          'FeedbackSubmitted',
          'Comments',
          'Attachments',
          'Created',
          'CreatedById',
          'Modified',
          'ModifiedById'
        )
        .expand('Interviewer', 'Coordinator')
        .top(5000)
        .orderBy('ScheduledDate', false);

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

        if (filter.interviewType && filter.interviewType.length > 0) {
          const typeFilters = filter.interviewType.map(t =>
            ValidationUtils.buildFilter('InterviewType', 'eq', t)
          );
          filters.push(`(${typeFilters.join(' or ')})`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.interviewerId) {
          ValidationUtils.validateInteger(filter.interviewerId, 'InterviewerId', 1);
          filters.push(`InterviewerId eq ${filter.interviewerId}`);
        }

        if (filter.result && filter.result.length > 0) {
          const resultFilters = filter.result.map(r =>
            ValidationUtils.buildFilter('Result', 'eq', r)
          );
          filters.push(`(${resultFilters.join(' or ')})`);
        }

        if (filter.fromDate) {
          ValidationUtils.validateDate(filter.fromDate, 'FromDate');
          filters.push(`ScheduledDate ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          ValidationUtils.validateDate(filter.toDate, 'ToDate');
          filters.push(`ScheduledDate le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query();
      return items.map(this.mapInterviewFromSP);
    } catch (error) {
      logger.error('InterviewService', 'Error getting interviews:', error);
      throw new Error(`Failed to retrieve interviews: ${error.message}`);
    }
  }

  /**
   * Get a single interview by ID
   */
  public async getInterviewById(id: number): Promise<IInterview> {
    try {
      ValidationUtils.validateInteger(id, 'Interview ID', 1);

      const item = await this.sp.web.lists.getByTitle(this.INTERVIEWS_LIST).items
        .getById(id)
        .select(
          'Id',
          'CandidateId',
          'JobRequisitionId',
          'InterviewType',
          'ScheduledDate',
          'Duration',
          'Status',
          'Location',
          'MeetingLink',
          'RoomNumber',
          'IsVirtual',
          'InterviewerId',
          'Interviewer/Title',
          'Interviewer/EMail',
          'PanelMembers',
          'CoordinatorId',
          'Coordinator/Title',
          'Result',
          'OverallScore',
          'TechnicalScore',
          'CommunicationScore',
          'ProblemSolvingScore',
          'CulturalFitScore',
          'Feedback',
          'Strengths',
          'Weaknesses',
          'Recommendation',
          'HiringRecommendation',
          'QuestionsAsked',
          'CandidateQuestions',
          'TechnicalAssessmentUrl',
          'PresentationUrl',
          'NextSteps',
          'FollowUpRequired',
          'FollowUpDate',
          'CandidateNotified',
          'ReminderSent',
          'FeedbackSubmitted',
          'Comments',
          'Attachments',
          'Created',
          'CreatedById',
          'Modified',
          'ModifiedById'
        )
        .expand('Interviewer', 'Coordinator')();

      return this.mapInterviewFromSP(item);
    } catch (error) {
      logger.error('InterviewService', `Error getting interview ${id}:`, error);
      throw new Error(`Failed to retrieve interview: ${error.message}`);
    }
  }

  /**
   * Get interviews for a specific candidate
   */
  public async getInterviewsForCandidate(candidateId: number): Promise<IInterview[]> {
    return this.getInterviews({ candidateId });
  }

  /**
   * Get upcoming interviews for an interviewer
   */
  public async getUpcomingInterviewsForInterviewer(interviewerId: number): Promise<IInterview[]> {
    const today = new Date();
    return this.getInterviews({
      interviewerId,
      fromDate: today,
      status: [InterviewStatus.Scheduled, InterviewStatus.Confirmed]
    });
  }

  /**
   * Create a new interview
   */
  public async createInterview(interview: Partial<IInterview>): Promise<number> {
    try {
      // Validation
      if (!interview.CandidateId || !interview.ScheduledDate) {
        throw new Error('CandidateId and ScheduledDate are required');
      }

      ValidationUtils.validateInteger(interview.CandidateId, 'CandidateId', 1);
      ValidationUtils.validateDate(interview.ScheduledDate, 'ScheduledDate');
      ValidationUtils.validateEnum(interview.InterviewType, InterviewType, 'InterviewType');
      ValidationUtils.validateEnum(interview.Status, InterviewStatus, 'Status');

      if (interview.Duration) {
        ValidationUtils.validateInteger(interview.Duration, 'Duration', 15, 480);
      }

      if (interview.InterviewerId) {
        ValidationUtils.validateInteger(interview.InterviewerId, 'InterviewerId', 1);

        // Check for scheduling conflicts
        const conflicts = await this.checkSchedulingConflicts(
          interview.InterviewerId,
          interview.ScheduledDate,
          interview.Duration || 60
        );

        if (conflicts.conflictingInterviews.length > 0) {
          logger.warn('InterviewService', 'Warning: Interviewer has ${conflicts.conflictingInterviews.length} conflicting interviews');
        }
      }

      // Validate scores if provided
      if (interview.OverallScore !== undefined) {
        ValidationUtils.validateInteger(interview.OverallScore, 'OverallScore', 1, 5);
      }
      if (interview.TechnicalScore !== undefined) {
        ValidationUtils.validateInteger(interview.TechnicalScore, 'TechnicalScore', 1, 5);
      }
      if (interview.CommunicationScore !== undefined) {
        ValidationUtils.validateInteger(interview.CommunicationScore, 'CommunicationScore', 1, 5);
      }
      if (interview.ProblemSolvingScore !== undefined) {
        ValidationUtils.validateInteger(interview.ProblemSolvingScore, 'ProblemSolvingScore', 1, 5);
      }
      if (interview.CulturalFitScore !== undefined) {
        ValidationUtils.validateInteger(interview.CulturalFitScore, 'CulturalFitScore', 1, 5);
      }

      // Prepare item data
      const itemData: any = {
        CandidateId: interview.CandidateId,
        JobRequisitionId: interview.JobRequisitionId || null,
        InterviewType: interview.InterviewType,
        ScheduledDate: interview.ScheduledDate,
        Duration: interview.Duration || 60,
        Status: interview.Status || InterviewStatus.Scheduled,
        Location: interview.Location ? ValidationUtils.sanitizeInput(interview.Location) : null,
        MeetingLink: interview.MeetingLink ? ValidationUtils.sanitizeInput(interview.MeetingLink) : null,
        RoomNumber: interview.RoomNumber ? ValidationUtils.sanitizeInput(interview.RoomNumber) : null,
        IsVirtual: interview.IsVirtual || false,
        PanelMembers: interview.PanelMembers || null,
        Result: interview.Result || null,
        OverallScore: interview.OverallScore || null,
        TechnicalScore: interview.TechnicalScore || null,
        CommunicationScore: interview.CommunicationScore || null,
        ProblemSolvingScore: interview.ProblemSolvingScore || null,
        CulturalFitScore: interview.CulturalFitScore || null,
        Feedback: interview.Feedback ? ValidationUtils.sanitizeHtml(interview.Feedback) : null,
        Strengths: interview.Strengths ? ValidationUtils.sanitizeHtml(interview.Strengths) : null,
        Weaknesses: interview.Weaknesses ? ValidationUtils.sanitizeHtml(interview.Weaknesses) : null,
        Recommendation: interview.Recommendation ? ValidationUtils.sanitizeHtml(interview.Recommendation) : null,
        HiringRecommendation: interview.HiringRecommendation || null,
        QuestionsAsked: interview.QuestionsAsked || null,
        CandidateQuestions: interview.CandidateQuestions ? ValidationUtils.sanitizeHtml(interview.CandidateQuestions) : null,
        TechnicalAssessmentUrl: interview.TechnicalAssessmentUrl ? ValidationUtils.sanitizeInput(interview.TechnicalAssessmentUrl) : null,
        PresentationUrl: interview.PresentationUrl ? ValidationUtils.sanitizeInput(interview.PresentationUrl) : null,
        NextSteps: interview.NextSteps ? ValidationUtils.sanitizeHtml(interview.NextSteps) : null,
        FollowUpRequired: interview.FollowUpRequired || false,
        FollowUpDate: interview.FollowUpDate || null,
        CandidateNotified: interview.CandidateNotified || false,
        ReminderSent: interview.ReminderSent || false,
        FeedbackSubmitted: interview.FeedbackSubmitted || false,
        Notes: interview.Notes ? ValidationUtils.sanitizeHtml(interview.Notes) : null,
        Attachments: interview.Attachments || null
      };

      // Add lookup fields
      if (interview.InterviewerId) {
        itemData.InterviewerId = interview.InterviewerId;
      }
      if (interview.CoordinatorId) {
        ValidationUtils.validateInteger(interview.CoordinatorId, 'CoordinatorId', 1);
        itemData.CoordinatorId = interview.CoordinatorId;
      }

      const result = await this.sp.web.lists.getByTitle(this.INTERVIEWS_LIST).items.add(itemData);

      // Log activity for candidate
      await this.logCandidateActivity(
        interview.CandidateId,
        'Interview',
        `${interview.InterviewType} scheduled`,
        `Interview scheduled for ${interview.ScheduledDate.toLocaleDateString()}`,
        result.data.Id
      );

      logger.debug('InterviewService', `Interview created successfully with ID: ${result.data.Id}`);
      return result.data.Id;
    } catch (error) {
      logger.error('InterviewService', 'Error creating interview:', error);
      throw new Error(`Failed to create interview: ${error.message}`);
    }
  }

  /**
   * Update an existing interview
   */
  public async updateInterview(id: number, updates: Partial<IInterview>): Promise<void> {
    try {
      ValidationUtils.validateInteger(id, 'Interview ID', 1);

      // Validate enums if provided
      if (updates.InterviewType) {
        ValidationUtils.validateEnum(updates.InterviewType, InterviewType, 'InterviewType');
      }
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, InterviewStatus, 'Status');
      }
      if (updates.Result) {
        ValidationUtils.validateEnum(updates.Result, InterviewResult, 'Result');
      }

      if (updates.ScheduledDate) {
        ValidationUtils.validateDate(updates.ScheduledDate, 'ScheduledDate');
      }

      if (updates.Duration !== undefined) {
        ValidationUtils.validateInteger(updates.Duration, 'Duration', 15, 480);
      }

      // Validate scores if provided
      if (updates.OverallScore !== undefined) {
        ValidationUtils.validateInteger(updates.OverallScore, 'OverallScore', 1, 5);
      }
      if (updates.TechnicalScore !== undefined) {
        ValidationUtils.validateInteger(updates.TechnicalScore, 'TechnicalScore', 1, 5);
      }
      if (updates.CommunicationScore !== undefined) {
        ValidationUtils.validateInteger(updates.CommunicationScore, 'CommunicationScore', 1, 5);
      }
      if (updates.ProblemSolvingScore !== undefined) {
        ValidationUtils.validateInteger(updates.ProblemSolvingScore, 'ProblemSolvingScore', 1, 5);
      }
      if (updates.CulturalFitScore !== undefined) {
        ValidationUtils.validateInteger(updates.CulturalFitScore, 'CulturalFitScore', 1, 5);
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
      if (updates.InterviewType) updateData.InterviewType = updates.InterviewType;
      if (updates.ScheduledDate) updateData.ScheduledDate = updates.ScheduledDate;
      if (updates.Duration !== undefined) updateData.Duration = updates.Duration;
      if (updates.Status) updateData.Status = updates.Status;
      if (updates.Location) updateData.Location = ValidationUtils.sanitizeInput(updates.Location);
      if (updates.MeetingLink) updateData.MeetingLink = ValidationUtils.sanitizeInput(updates.MeetingLink);
      if (updates.RoomNumber) updateData.RoomNumber = ValidationUtils.sanitizeInput(updates.RoomNumber);
      if (updates.IsVirtual !== undefined) updateData.IsVirtual = updates.IsVirtual;
      if (updates.PanelMembers) updateData.PanelMembers = updates.PanelMembers;
      if (updates.Result) updateData.Result = updates.Result;
      if (updates.OverallScore !== undefined) updateData.OverallScore = updates.OverallScore;
      if (updates.TechnicalScore !== undefined) updateData.TechnicalScore = updates.TechnicalScore;
      if (updates.CommunicationScore !== undefined) updateData.CommunicationScore = updates.CommunicationScore;
      if (updates.ProblemSolvingScore !== undefined) updateData.ProblemSolvingScore = updates.ProblemSolvingScore;
      if (updates.CulturalFitScore !== undefined) updateData.CulturalFitScore = updates.CulturalFitScore;
      if (updates.Feedback) updateData.Feedback = ValidationUtils.sanitizeHtml(updates.Feedback);
      if (updates.Strengths) updateData.Strengths = ValidationUtils.sanitizeHtml(updates.Strengths);
      if (updates.Weaknesses) updateData.Weaknesses = ValidationUtils.sanitizeHtml(updates.Weaknesses);
      if (updates.Recommendation) updateData.Recommendation = ValidationUtils.sanitizeHtml(updates.Recommendation);
      if (updates.HiringRecommendation) updateData.HiringRecommendation = updates.HiringRecommendation;
      if (updates.QuestionsAsked) updateData.QuestionsAsked = updates.QuestionsAsked;
      if (updates.CandidateQuestions) updateData.CandidateQuestions = ValidationUtils.sanitizeHtml(updates.CandidateQuestions);
      if (updates.TechnicalAssessmentUrl) updateData.TechnicalAssessmentUrl = ValidationUtils.sanitizeInput(updates.TechnicalAssessmentUrl);
      if (updates.PresentationUrl) updateData.PresentationUrl = ValidationUtils.sanitizeInput(updates.PresentationUrl);
      if (updates.NextSteps) updateData.NextSteps = ValidationUtils.sanitizeHtml(updates.NextSteps);
      if (updates.FollowUpRequired !== undefined) updateData.FollowUpRequired = updates.FollowUpRequired;
      if (updates.FollowUpDate) updateData.FollowUpDate = updates.FollowUpDate;
      if (updates.CandidateNotified !== undefined) updateData.CandidateNotified = updates.CandidateNotified;
      if (updates.ReminderSent !== undefined) updateData.ReminderSent = updates.ReminderSent;
      if (updates.FeedbackSubmitted !== undefined) updateData.FeedbackSubmitted = updates.FeedbackSubmitted;
      if (updates.Notes) updateData.Notes = ValidationUtils.sanitizeHtml(updates.Notes);
      if (updates.Attachments) updateData.Attachments = updates.Attachments;

      // Update lookup fields
      if (updates.InterviewerId) {
        ValidationUtils.validateInteger(updates.InterviewerId, 'InterviewerId', 1);
        updateData.InterviewerId = updates.InterviewerId;
      }
      if (updates.CoordinatorId) {
        ValidationUtils.validateInteger(updates.CoordinatorId, 'CoordinatorId', 1);
        updateData.CoordinatorId = updates.CoordinatorId;
      }

      await this.sp.web.lists.getByTitle(this.INTERVIEWS_LIST).items.getById(id).update(updateData);

      logger.debug('InterviewService', `Interview ${id} updated successfully`);
    } catch (error) {
      logger.error('InterviewService', `Error updating interview ${id}:`, error);
      throw new Error(`Failed to update interview: ${error.message}`);
    }
  }

  /**
   * Delete an interview
   */
  public async deleteInterview(id: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(id, 'Interview ID', 1);

      await this.sp.web.lists.getByTitle(this.INTERVIEWS_LIST).items.getById(id).delete();

      logger.debug('InterviewService', `Interview ${id} deleted successfully`);
    } catch (error) {
      logger.error('InterviewService', `Error deleting interview ${id}:`, error);
      throw new Error(`Failed to delete interview: ${error.message}`);
    }
  }

  // ==================== Interview Feedback ====================

  /**
   * Submit interview feedback
   */
  public async submitInterviewFeedback(
    interviewId: number,
    feedback: {
      result: InterviewResult;
      overallScore: number;
      technicalScore?: number;
      communicationScore?: number;
      problemSolvingScore?: number;
      culturalFitScore?: number;
      feedback: string;
      strengths?: string;
      weaknesses?: string;
      recommendation?: string;
      hiringRecommendation?: 'Strong Yes' | 'Yes' | 'Maybe' | 'No' | 'Strong No';
    }
  ): Promise<void> {
    try {
      ValidationUtils.validateInteger(interviewId, 'Interview ID', 1);
      ValidationUtils.validateEnum(feedback.result, InterviewResult, 'Result');
      ValidationUtils.validateInteger(feedback.overallScore, 'OverallScore', 1, 5);

      if (!feedback.feedback) {
        throw new Error('Feedback is required');
      }

      const interview = await this.getInterviewById(interviewId);

      await this.updateInterview(interviewId, {
        Status: InterviewStatus.Completed,
        Result: feedback.result,
        OverallScore: feedback.overallScore,
        TechnicalScore: feedback.technicalScore,
        CommunicationScore: feedback.communicationScore,
        ProblemSolvingScore: feedback.problemSolvingScore,
        CulturalFitScore: feedback.culturalFitScore,
        Feedback: feedback.feedback,
        Strengths: feedback.strengths,
        Weaknesses: feedback.weaknesses,
        Recommendation: feedback.recommendation,
        HiringRecommendation: feedback.hiringRecommendation,
        FeedbackSubmitted: true
      });

      // Log activity
      await this.logCandidateActivity(
        interview.CandidateId,
        'Interview',
        'Interview feedback submitted',
        `${interview.InterviewType} completed with result: ${feedback.result}`,
        interviewId
      );

      logger.debug('InterviewService', `Feedback submitted for interview ${interviewId}`);
    } catch (error) {
      logger.error('InterviewService', `Error submitting feedback for interview ${interviewId}:`, error);
      throw new Error(`Failed to submit feedback: ${error.message}`);
    }
  }

  // ==================== Scheduling ====================

  /**
   * Check for scheduling conflicts for an interviewer
   */
  public async checkSchedulingConflicts(
    interviewerId: number,
    scheduledDate: Date,
    duration: number
  ): Promise<IScheduleConflict> {
    try {
      ValidationUtils.validateInteger(interviewerId, 'InterviewerId', 1);
      ValidationUtils.validateDate(scheduledDate, 'ScheduledDate');
      ValidationUtils.validateInteger(duration, 'Duration', 15, 480);

      // Get all interviews for the interviewer on the same day
      const startOfDay = new Date(scheduledDate);
      startOfDay.setHours(0, 0, 0, 0);

      const endOfDay = new Date(scheduledDate);
      endOfDay.setHours(23, 59, 59, 999);

      const existingInterviews = await this.getInterviews({
        interviewerId,
        fromDate: startOfDay,
        toDate: endOfDay,
        status: [InterviewStatus.Scheduled, InterviewStatus.Confirmed]
      });

      // Check for time overlaps
      const proposedStart = scheduledDate.getTime();
      const proposedEnd = proposedStart + (duration * 60 * 1000);

      const conflictingInterviews = existingInterviews.filter(interview => {
        const existingStart = new Date(interview.ScheduledDate).getTime();
        const existingEnd = existingStart + (interview.Duration * 60 * 1000);

        // Check if times overlap
        return (proposedStart < existingEnd && proposedEnd > existingStart);
      });

      // Suggest alternative times (on the hour, 9 AM to 5 PM)
      const suggestedTimes: Date[] = [];
      if (conflictingInterviews.length > 0) {
        for (let hour = 9; hour <= 17; hour++) {
          const suggestionTime = new Date(scheduledDate);
          suggestionTime.setHours(hour, 0, 0, 0);

          const suggestionStart = suggestionTime.getTime();
          const suggestionEnd = suggestionStart + (duration * 60 * 1000);

          // Check if this time slot is free
          const hasConflict = existingInterviews.some(interview => {
            const existingStart = new Date(interview.ScheduledDate).getTime();
            const existingEnd = existingStart + (interview.Duration * 60 * 1000);
            return (suggestionStart < existingEnd && suggestionEnd > existingStart);
          });

          if (!hasConflict) {
            suggestedTimes.push(suggestionTime);
          }

          if (suggestedTimes.length >= 5) break; // Limit to 5 suggestions
        }
      }

      return {
        interviewerId,
        conflictingInterviews,
        suggestedTimes
      };
    } catch (error) {
      logger.error('InterviewService', 'Error checking scheduling conflicts:', error);
      throw new Error(`Failed to check scheduling conflicts: ${error.message}`);
    }
  }

  /**
   * Reschedule an interview
   */
  public async rescheduleInterview(
    interviewId: number,
    newDate: Date,
    reason?: string
  ): Promise<void> {
    try {
      ValidationUtils.validateInteger(interviewId, 'Interview ID', 1);
      ValidationUtils.validateDate(newDate, 'NewDate');

      const interview = await this.getInterviewById(interviewId);

      // Check for conflicts
      if (interview.InterviewerId) {
        const conflicts = await this.checkSchedulingConflicts(
          interview.InterviewerId,
          newDate,
          interview.Duration
        );

        if (conflicts.conflictingInterviews.length > 0) {
          throw new Error(`Scheduling conflict detected. Suggested times: ${conflicts.suggestedTimes.map(t => t.toLocaleString()).join(', ')}`);
        }
      }

      await this.updateInterview(interviewId, {
        ScheduledDate: newDate,
        Status: InterviewStatus.Rescheduled,
        Notes: reason ? `Rescheduled: ${ValidationUtils.sanitizeInput(reason)}` : 'Interview rescheduled'
      });

      // Log activity
      await this.logCandidateActivity(
        interview.CandidateId,
        'Interview',
        'Interview rescheduled',
        `${interview.InterviewType} rescheduled to ${newDate.toLocaleDateString()}${reason ? ': ' + reason : ''}`,
        interviewId
      );

      logger.debug('InterviewService', `Interview ${interviewId} rescheduled to ${newDate}`);
    } catch (error) {
      logger.error('InterviewService', `Error rescheduling interview ${interviewId}:`, error);
      throw new Error(`Failed to reschedule interview: ${error.message}`);
    }
  }

  /**
   * Cancel an interview
   */
  public async cancelInterview(interviewId: number, reason?: string): Promise<void> {
    try {
      ValidationUtils.validateInteger(interviewId, 'Interview ID', 1);

      const interview = await this.getInterviewById(interviewId);

      await this.updateInterview(interviewId, {
        Status: InterviewStatus.Cancelled,
        Notes: reason ? `Cancelled: ${ValidationUtils.sanitizeInput(reason)}` : 'Interview cancelled'
      });

      // Log activity
      await this.logCandidateActivity(
        interview.CandidateId,
        'Interview',
        'Interview cancelled',
        `${interview.InterviewType} cancelled${reason ? ': ' + reason : ''}`,
        interviewId
      );

      logger.debug('InterviewService', `Interview ${interviewId} cancelled`);
    } catch (error) {
      logger.error('InterviewService', `Error cancelling interview ${interviewId}:`, error);
      throw new Error(`Failed to cancel interview: ${error.message}`);
    }
  }

  /**
   * Mark interview as no-show
   */
  public async markAsNoShow(interviewId: number): Promise<void> {
    try {
      ValidationUtils.validateInteger(interviewId, 'Interview ID', 1);

      const interview = await this.getInterviewById(interviewId);

      await this.updateInterview(interviewId, {
        Status: InterviewStatus.NoShow
      });

      // Log activity
      await this.logCandidateActivity(
        interview.CandidateId,
        'Interview',
        'Interview no-show',
        `Candidate did not attend ${interview.InterviewType}`,
        interviewId
      );

      logger.debug('InterviewService', `Interview ${interviewId} marked as no-show`);
    } catch (error) {
      logger.error('InterviewService', `Error marking interview ${interviewId} as no-show:`, error);
      throw new Error(`Failed to mark as no-show: ${error.message}`);
    }
  }

  // ==================== Analytics ====================

  /**
   * Get average interview scores for a candidate
   */
  public async getAverageScoresForCandidate(candidateId: number): Promise<{
    overallAverage: number;
    technicalAverage: number;
    communicationAverage: number;
    problemSolvingAverage: number;
    culturalFitAverage: number;
    interviewCount: number;
  }> {
    try {
      ValidationUtils.validateInteger(candidateId, 'Candidate ID', 1);

      const interviews = await this.getInterviewsForCandidate(candidateId);
      const completedInterviews = interviews.filter(i => i.FeedbackSubmitted && i.OverallScore);

      if (completedInterviews.length === 0) {
        return {
          overallAverage: 0,
          technicalAverage: 0,
          communicationAverage: 0,
          problemSolvingAverage: 0,
          culturalFitAverage: 0,
          interviewCount: 0
        };
      }

      const calculateAverage = (field: keyof IInterview): number => {
        const scores = completedInterviews
          .map(i => i[field] as number)
          .filter(score => score && score > 0);
        return scores.length > 0
          ? scores.reduce((sum, score) => sum + score, 0) / scores.length
          : 0;
      };

      return {
        overallAverage: calculateAverage('OverallScore'),
        technicalAverage: calculateAverage('TechnicalScore'),
        communicationAverage: calculateAverage('CommunicationScore'),
        problemSolvingAverage: calculateAverage('ProblemSolvingScore'),
        culturalFitAverage: calculateAverage('CulturalFitScore'),
        interviewCount: completedInterviews.length
      };
    } catch (error) {
      logger.error('InterviewService', `Error calculating average scores for candidate ${candidateId}:`, error);
      throw new Error(`Failed to calculate average scores: ${error.message}`);
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
    description?: string,
    relatedInterviewId?: number
  ): Promise<number> {
    try {
      const activityData: any = {
        CandidateId: candidateId,
        ActivityType: activityType,
        ActivityDate: new Date(),
        Subject: subject ? ValidationUtils.sanitizeInput(subject) : null,
        Description: description ? ValidationUtils.sanitizeHtml(description) : null,
        RelatedInterviewId: relatedInterviewId || null
      };

      const result = await this.sp.web.lists.getByTitle(this.CANDIDATE_ACTIVITIES_LIST).items.add(activityData);
      return result.data.Id;
    } catch (error) {
      logger.error('InterviewService', 'Error logging candidate activity:', error);
      // Don't throw - activity logging should not block main operations
      return 0;
    }
  }

  /**
   * Map SharePoint item to IInterview
   */
  private mapInterviewFromSP(item: any): IInterview {
    return {
      Id: item.Id,
      CandidateId: item.CandidateId,
      JobRequisitionId: item.JobRequisitionId,
      InterviewType: item.InterviewType,
      ScheduledDate: new Date(item.ScheduledDate),
      Duration: item.Duration,
      Status: item.Status,
      Location: item.Location,
      MeetingLink: item.MeetingLink,
      RoomNumber: item.RoomNumber,
      IsVirtual: item.IsVirtual,
      InterviewerId: item.InterviewerId,
      Interviewer: item.Interviewer,
      PanelMembers: item.PanelMembers,
      CoordinatorId: item.CoordinatorId,
      Coordinator: item.Coordinator,
      Result: item.Result,
      OverallScore: item.OverallScore,
      TechnicalScore: item.TechnicalScore,
      CommunicationScore: item.CommunicationScore,
      ProblemSolvingScore: item.ProblemSolvingScore,
      CulturalFitScore: item.CulturalFitScore,
      Feedback: item.Feedback,
      Strengths: item.Strengths,
      Weaknesses: item.Weaknesses,
      Recommendation: item.Recommendation,
      HiringRecommendation: item.HiringRecommendation,
      QuestionsAsked: item.QuestionsAsked,
      CandidateQuestions: item.CandidateQuestions,
      TechnicalAssessmentUrl: item.TechnicalAssessmentUrl,
      PresentationUrl: item.PresentationUrl,
      NextSteps: item.NextSteps,
      FollowUpRequired: item.FollowUpRequired,
      FollowUpDate: item.FollowUpDate ? new Date(item.FollowUpDate) : undefined,
      CandidateNotified: item.CandidateNotified,
      ReminderSent: item.ReminderSent,
      FeedbackSubmitted: item.FeedbackSubmitted,
      Notes: item.Notes,
      Attachments: item.Attachments,
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedById: item.CreatedById,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      ModifiedById: item.ModifiedById
    };
  }
}
