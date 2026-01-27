// CV Management Models
// Interfaces for managing candidate CVs/resumes

import { IBaseListItem, IUser } from './ICommon';

/**
 * CV Status
 */
export enum CVStatus {
  New = 'New',
  Screening = 'Screening',
  Qualified = 'Qualified',
  Shortlisted = 'Shortlisted',
  Rejected = 'Rejected',
  InInterview = 'In Interview',
  Offered = 'Offered',
  Hired = 'Hired',
  Archived = 'Archived'
}

/**
 * CV Source
 */
export enum CVSource {
  DirectApplication = 'Direct Application',
  LinkedIn = 'LinkedIn',
  JobBoard = 'Job Board',
  Referral = 'Referral',
  Recruiter = 'Recruiter',
  CareerFair = 'Career Fair',
  Email = 'Email',
  Other = 'Other'
}

/**
 * Experience Level
 */
export enum ExperienceLevel {
  EntryLevel = 'Entry Level',
  Junior = '1-3 Years',
  MidLevel = '3-5 Years',
  Senior = '5-10 Years',
  Expert = '10+ Years',
  Executive = 'Executive'
}

/**
 * Education Level
 */
export enum EducationLevel {
  HighSchool = 'High School',
  Associate = 'Associate Degree',
  Bachelor = 'Bachelor Degree',
  Master = 'Master Degree',
  Doctorate = 'Doctorate',
  Professional = 'Professional Certificate'
}

/**
 * CV/Resume Record
 */
export interface ICV extends IBaseListItem {
  // Candidate Information
  CandidateName: string;
  Email: string;
  Phone?: string;
  LinkedInProfile?: string;
  Location?: string;

  // CV Details
  CVFileName: string;
  CVFileUrl?: string; // SharePoint document URL
  CVFileSize?: number; // in KB
  SubmissionDate: Date;
  Source: CVSource;

  // Position Information
  PositionAppliedFor?: string;
  JobRequisitionId?: number;
  Department?: string;

  // Qualifications
  YearsOfExperience?: number;
  ExperienceLevel?: ExperienceLevel;
  HighestEducation?: EducationLevel;
  Skills?: string[]; // Array of skills
  Certifications?: string;
  Languages?: string;

  // Status & Tracking
  Status: CVStatus;
  QualificationScore?: number; // 0-100
  ScreeningNotes?: string;
  Reviewer?: IUser;
  ReviewDate?: Date;

  // Shortlisting
  IsShortlisted: boolean;
  ShortlistedBy?: IUser;
  ShortlistedDate?: Date;
  ShortlistReason?: string;

  // Rejection
  RejectionReason?: string;
  RejectedBy?: IUser;
  RejectionDate?: Date;

  // Search & Matching
  KeywordTags?: string[]; // For search functionality
  MatchScore?: number; // 0-100 match with job requirements

  // Follow-up
  NextAction?: string;
  NextActionDate?: Date;
  InterviewScheduled?: boolean;
  InterviewDate?: Date;

  // Additional Info
  SalaryExpectation?: string;
  NoticePeriod?: string;
  Availability?: string;
  Notes?: string;
}

/**
 * CV Qualification Criteria
 */
export interface ICVQualificationCriteria extends IBaseListItem {
  JobRequisitionId?: number;
  PositionTitle: string;
  Department?: string;

  // Required Criteria
  MinimumEducation: EducationLevel;
  MinimumExperience: number; // in years
  RequiredSkills: string[];
  RequiredCertifications?: string[];

  // Preferred Criteria
  PreferredSkills?: string[];
  PreferredCertifications?: string[];
  PreferredEducation?: EducationLevel;

  // Scoring Weights
  EducationWeight: number; // percentage
  ExperienceWeight: number;
  SkillsWeight: number;
  CertificationsWeight: number;

  // Auto-qualification Rules
  AutoQualifyScore?: number; // Score threshold for auto-qualification
  AutoRejectScore?: number; // Score threshold for auto-rejection

  IsActive: boolean;
}

/**
 * CV Search Filters
 */
export interface ICVSearchFilters {
  keyword?: string;
  status?: CVStatus[];
  source?: CVSource[];
  positionAppliedFor?: string;
  department?: string;
  experienceLevel?: ExperienceLevel[];
  education?: EducationLevel[];
  skills?: string[];
  minYearsExperience?: number;
  maxYearsExperience?: number;
  submissionDateFrom?: Date;
  submissionDateTo?: Date;
  isShortlisted?: boolean;
  minQualificationScore?: number;
  reviewer?: string;
  location?: string;
}

/**
 * CV Bulk Action
 */
export interface ICVBulkAction {
  cvIds: number[];
  action: 'qualify' | 'shortlist' | 'reject' | 'archive' | 'assign' | 'tag';
  reason?: string;
  assignTo?: string; // User email
  tags?: string[];
  newStatus?: CVStatus;
}

/**
 * CV Analytics
 */
export interface ICVAnalytics {
  totalCVs: number;
  newCVs: number;
  qualifiedCVs: number;
  shortlistedCVs: number;
  rejectedCVs: number;
  averageQualificationScore: number;
  cvsBySource: { [key: string]: number };
  cvsByStatus: { [key: string]: number };
  cvsByPosition: { [key: string]: number };
  topSkills: { skill: string; count: number }[];
  averageResponseTime: number; // in days
  conversionRate: number; // percentage from submitted to hired
}

/**
 * CV Review History
 */
export interface ICVReviewHistory extends IBaseListItem {
  CVId: number;
  CandidateName: string;
  ReviewedBy: IUser;
  ReviewDate: Date;
  PreviousStatus: CVStatus;
  NewStatus: CVStatus;
  Action: string;
  Comments?: string;
  QualificationScore?: number;
}

/**
 * CV Attachment
 */
export interface ICVAttachment {
  fileName: string;
  fileUrl: string;
  fileType: string; // 'CV', 'Cover Letter', 'Certificate', 'Portfolio', 'Other'
  fileSize: number; // in KB
  uploadedDate: Date;
  uploadedBy?: IUser;
}

/**
 * CV Email Template
 */
export interface ICVEmailTemplate extends IBaseListItem {
  TemplateName: string;
  TemplateType: 'Acknowledgment' | 'Rejection' | 'Shortlist' | 'Interview Invitation' | 'Follow-up';
  Subject: string;
  Body: string; // HTML content with placeholders
  IsActive: boolean;
  IsDefault: boolean;
}

/**
 * CV Import/Export
 */
export interface ICVImportData {
  candidateName: string;
  email: string;
  phone?: string;
  positionAppliedFor?: string;
  source?: string;
  yearsOfExperience?: number;
  education?: string;
  skills?: string;
  notes?: string;
}

export interface ICVExportOptions {
  format: 'Excel' | 'CSV' | 'PDF';
  filters?: ICVSearchFilters;
  includeFields: string[];
  includeAttachments: boolean;
}

/**
 * CV Duplicate Check
 */
export interface ICVDuplicateCheck {
  email?: string;
  phone?: string;
  linkedInProfile?: string;
}

export interface ICVDuplicateResult {
  isDuplicate: boolean;
  matchType: 'email' | 'phone' | 'linkedin' | 'none';
  existingCV?: ICV;
  message?: string;
}

/**
 * CV Screening Question
 */
export interface ICVScreeningQuestion extends IBaseListItem {
  JobRequisitionId?: number;
  Question: string;
  QuestionType: 'Text' | 'Number' | 'Boolean' | 'MultipleChoice' | 'File';
  Options?: string[]; // For multiple choice
  IsRequired: boolean;
  Weight: number; // For scoring
  CorrectAnswer?: string; // For auto-scoring
  DisplayOrder: number;
  IsActive: boolean;
}

/**
 * CV Screening Answer
 */
export interface ICVScreeningAnswer extends IBaseListItem {
  CVId: number;
  QuestionId: number;
  Question: string;
  Answer: string;
  Score?: number;
  AnsweredDate: Date;
}
