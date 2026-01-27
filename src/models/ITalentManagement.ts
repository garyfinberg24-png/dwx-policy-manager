// Talent Management & Recruitment Models
// Comprehensive HR recruitment, candidate tracking, and hiring workflow

export enum CandidateStatus {
  New = 'New',
  Screening = 'Screening',
  PhoneScreen = 'Phone Screen',
  Interview = 'Interview',
  TechnicalAssessment = 'Technical Assessment',
  BackgroundCheck = 'Background Check',
  Offer = 'Offer',
  OfferAccepted = 'Offer Accepted',
  OfferDeclined = 'Offer Declined',
  Hired = 'Hired',
  Rejected = 'Rejected',
  Withdrawn = 'Withdrawn',
  OnHold = 'On Hold'
}

export enum InterviewType {
  PhoneScreen = 'Phone Screen',
  VideoInterview = 'Video Interview',
  InPerson = 'In-Person',
  Panel = 'Panel Interview',
  Technical = 'Technical Interview',
  Behavioral = 'Behavioral Interview',
  CaseStudy = 'Case Study',
  PresentationDemo = 'Presentation/Demo',
  FinalRound = 'Final Round'
}

export enum InterviewStatus {
  Scheduled = 'Scheduled',
  Confirmed = 'Confirmed',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Rescheduled = 'Rescheduled',
  Cancelled = 'Cancelled',
  NoShow = 'No Show'
}

export enum InterviewResult {
  StrongHire = 'Strong Hire',
  Hire = 'Hire',
  MaybeHire = 'Maybe Hire',
  NoHire = 'No Hire',
  StrongNoHire = 'Strong No Hire',
  Pending = 'Pending'
}

export enum JobRequisitionStatus {
  Draft = 'Draft',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Open = 'Open',
  OnHold = 'On Hold',
  Filled = 'Filled',
  Cancelled = 'Cancelled',
  Closed = 'Closed'
}

export enum ApplicationSource {
  CompanyWebsite = 'Company Website',
  LinkedIn = 'LinkedIn',
  Indeed = 'Indeed',
  Glassdoor = 'Glassdoor',
  Referral = 'Employee Referral',
  Recruiter = 'Recruiter',
  JobBoard = 'Job Board',
  CareerFair = 'Career Fair',
  DirectApply = 'Direct Apply',
  Other = 'Other'
}

export enum OfferStatus {
  Draft = 'Draft',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Sent = 'Sent',
  Accepted = 'Accepted',
  Declined = 'Declined',
  Withdrawn = 'Withdrawn',
  Expired = 'Expired'
}

export enum EmploymentType {
  FullTime = 'Full-Time',
  PartTime = 'Part-Time',
  Contract = 'Contract',
  Temporary = 'Temporary',
  Intern = 'Intern',
  Consultant = 'Consultant'
}

export enum Priority {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

// ==================== Job Requisition ====================

export interface IJobRequisition {
  Id?: number;
  Title: string;
  JobCode?: string;
  Department: string;
  Location: string;
  EmploymentType: EmploymentType;
  Status: JobRequisitionStatus;

  // Reporting
  HiringManagerId?: number;
  HiringManager?: any; // User lookup
  RecruiterId?: number;
  Recruiter?: any; // User lookup

  // Position Details
  NumberOfOpenings: number;
  FilledPositions?: number;
  Priority: Priority;
  TargetStartDate?: Date;
  TargetFillDate?: Date;

  // Job Description
  JobDescription: string;
  Responsibilities?: string;
  Requirements?: string;
  PreferredQualifications?: string;
  Skills?: string; // JSON array
  Certifications?: string;

  // Compensation
  SalaryRangeMin?: number;
  SalaryRangeMax?: number;
  Currency?: string;
  BonusEligible?: boolean;
  BenefitsDescription?: string;

  // Approval
  ApprovalRequired?: boolean;
  ApprovedById?: number;
  ApprovedBy?: any;
  ApprovalDate?: Date;
  ApprovalNotes?: string;

  // Budget
  BudgetCode?: string;
  EstimatedTotalCost?: number;

  // Posting
  IsPublished?: boolean;
  PostedDate?: Date;
  PostingExpiration?: Date;
  ExternalJobBoards?: string; // JSON array

  // Metadata
  Notes?: string;
  Attachments?: string; // JSON array
  Created?: Date;
  CreatedById?: number;
  Modified?: Date;
  ModifiedById?: number;
}

// ==================== Candidate ====================

export interface ICandidate {
  Id?: number;
  FirstName: string;
  LastName: string;
  Email: string;
  Phone?: string;
  LinkedInURL?: string;

  // Current Position
  CurrentCompany?: string;
  CurrentPosition?: string;
  YearsOfExperience?: number;

  // Education
  HighestDegree?: string;

  // Location
  City?: string;
  State?: string;
  Country?: string;
  IsWillingToRelocate?: boolean;

  // Application
  Status: CandidateStatus;
  ApplicationDate: Date;
  ApplicationSource: ApplicationSource;
  ReferredById?: number;
  ReferredBy?: any; // User lookup
  ReferralBonus?: number;

  // Applied Position
  JobRequisitionId?: number;
  JobRequisition?: IJobRequisition;

  // Documents
  ResumeURL?: string;
  CoverLetterURL?: string;
  AdditionalDocuments?: string; // JSON array

  // Assessment
  OverallRating?: number; // 1-5 stars
  TechnicalSkillsRating?: number;
  CommunicationSkillsRating?: number;
  CulturalFitRating?: number;

  // Skills
  Skills?: string; // JSON array
  Certifications?: string; // JSON array
  Education?: string; // JSON array
  Languages?: string; // JSON array

  // Compensation Expectations
  ExpectedSalary?: number;
  CurrentSalary?: number;
  NoticePeriod?: number; // Days
  EarliestStartDate?: Date;

  // Background Check
  BackgroundCheckStatus?: 'Not Started' | 'In Progress' | 'Completed' | 'Failed';
  BackgroundCheckDate?: Date;
  BackgroundCheckNotes?: string;

  // Offer
  OfferId?: number;
  OfferAmount?: number;
  OfferDate?: Date;

  // Rejection
  RejectionReason?: string;
  RejectionDate?: Date;
  IsEligibleForRehire?: boolean;

  // Communication
  LastContactDate?: Date;
  NextFollowUpDate?: Date;

  // Diversity & Inclusion (optional, self-reported)
  Gender?: string;
  Ethnicity?: string;
  Veteran?: boolean;
  Disability?: boolean;

  // Metadata
  Notes?: string;
  Tags?: string; // JSON array for easy filtering
  Created?: Date;
  CreatedById?: number;
  Modified?: Date;
  ModifiedById?: number;
}

// ==================== Interview ====================

export interface IInterview {
  Id?: number;
  CandidateId: number;
  Candidate?: ICandidate;
  JobRequisitionId?: number;
  JobRequisition?: IJobRequisition;

  // Schedule
  InterviewType: InterviewType;
  ScheduledDate: Date;
  Duration: number; // Minutes
  Status: InterviewStatus;

  // Location/Method
  Location?: string;
  MeetingLink?: string; // Teams/Zoom link
  RoomNumber?: string;
  IsVirtual?: boolean;

  // Participants
  InterviewerId?: number;
  Interviewer?: any; // User lookup
  PanelMembers?: string; // JSON array of user IDs
  CoordinatorId?: number;
  Coordinator?: any; // User lookup

  // Evaluation
  Result?: InterviewResult;
  OverallScore?: number; // 1-5
  TechnicalScore?: number;
  CommunicationScore?: number;
  ProblemSolvingScore?: number;
  CulturalFitScore?: number;

  // Feedback
  Feedback?: string;
  Strengths?: string;
  Weaknesses?: string;
  Recommendation?: string;
  HiringRecommendation?: 'Strong Yes' | 'Yes' | 'Maybe' | 'No' | 'Strong No';

  // Questions & Notes
  QuestionsAsked?: string; // JSON array
  CandidateQuestions?: string;
  TechnicalAssessmentUrl?: string;
  PresentationUrl?: string;

  // Follow-up
  NextSteps?: string;
  FollowUpRequired?: boolean;
  FollowUpDate?: Date;

  // Notifications
  CandidateNotified?: boolean;
  ReminderSent?: boolean;
  FeedbackSubmitted?: boolean;

  // Metadata
  Notes?: string;
  Attachments?: string; // JSON array
  Created?: Date;
  CreatedById?: number;
  Modified?: Date;
  ModifiedById?: number;
}

// ==================== Job Offer ====================

export interface IJobOffer {
  Id?: number;
  CandidateId: number;
  Candidate?: ICandidate;
  JobRequisitionId: number;
  JobRequisition?: IJobRequisition;

  // Offer Details
  JobTitle: string;
  Department: string;
  Location: string;
  EmploymentType: EmploymentType;
  StartDate?: Date;

  // Compensation
  BaseSalary: number;
  Currency?: string;
  BonusAmount?: number;
  SigningBonus?: number;
  Equity?: string;
  Benefits?: string; // JSON array

  // Work Arrangement
  IsRemote?: boolean;
  HybridSchedule?: string;
  WorkingHours?: string;

  // Reporting
  ReportsToId?: number;
  ReportsTo?: any; // User lookup

  // Offer Status
  Status: OfferStatus;
  OfferDate?: Date;
  ExpirationDate?: Date;
  AcceptanceDate?: Date;
  DeclineReason?: string;

  // Approval Workflow
  ApprovalRequired?: boolean;
  ApprovedById?: number;
  ApprovedBy?: any;
  ApprovalDate?: Date;
  ApprovalNotes?: string;

  // Contract
  OfferLetterUrl?: string;
  ContractUrl?: string;
  IsSigned?: boolean;
  SignedDate?: Date;

  // Onboarding
  OnboardingProcessId?: number; // Link to JML Process
  OnboardingStartDate?: Date;

  // Metadata
  Notes?: string;
  Attachments?: string; // JSON array
  Created?: Date;
  CreatedById?: number;
  Modified?: Date;
  ModifiedById?: number;
}

// ==================== Candidate Activity ====================

export interface ICandidateActivity {
  Id?: number;
  CandidateId: number;
  ActivityType: 'Email' | 'Phone Call' | 'Meeting' | 'Interview' | 'Note' | 'Status Change' | 'Document Upload';
  ActivityDate: Date;
  Subject?: string;
  Description?: string;
  PerformedById?: number;
  PerformedBy?: any;
  RelatedInterviewId?: number;
  Attachments?: string;
  Created?: Date;
}

// ==================== Interview Feedback Template ====================

export interface IInterviewFeedbackTemplate {
  Id?: number;
  Title: string;
  InterviewType: InterviewType;
  IsActive: boolean;
  Questions?: string; // JSON array
  EvaluationCriteria?: string; // JSON array
  ScoringRubric?: string;
  Created?: Date;
  Modified?: Date;
}

// ==================== Talent Pool ====================

export interface ITalentPool {
  Id?: number;
  Title: string;
  Description?: string;
  CandidateIds?: string; // JSON array
  TargetRoles?: string; // JSON array
  CreatedById?: number;
  IsActive: boolean;
  Created?: Date;
  Modified?: Date;
}

// ==================== Recruitment Metrics ====================

export interface IRecruitmentMetrics {
  // Pipeline
  totalCandidates: number;
  newCandidates: number;
  candidatesByStatus: {
    [key in CandidateStatus]?: number;
  };

  // Job Requisitions
  totalRequisitions: number;
  openRequisitions: number;
  filledRequisitions: number;
  avgTimeToFill: number; // Days
  avgTimeToHire: number; // Days

  // Interviews
  totalInterviews: number;
  upcomingInterviews: number;
  completedInterviews: number;
  interviewCompletionRate: number; // Percentage

  // Offers
  totalOffers: number;
  pendingOffers: number;
  acceptedOffers: number;
  declinedOffers: number;
  offerAcceptanceRate: number; // Percentage

  // Sources
  candidatesBySource: {
    [key in ApplicationSource]?: number;
  };
  bestPerformingSource?: ApplicationSource;

  // Quality
  avgCandidateRating: number;
  avgInterviewScore: number;

  // Cost
  totalRecruitmentCost: number;
  costPerHire: number;
  referralBonusesPaid: number;

  // Recent Activity
  recentActivity: {
    newApplications: number;
    scheduledInterviews: number;
    completedInterviews: number;
    offersExtended: number;
  };
}

// ==================== Filter Criteria ====================

export interface ICandidateFilterCriteria {
  status?: CandidateStatus[];
  jobRequisitionId?: number;
  source?: ApplicationSource[];
  minRating?: number;
  skills?: string[];
  location?: string;
  minExperience?: number;
  maxExperience?: number;
  searchTerm?: string;
  fromDate?: Date;
  toDate?: Date;
}

export interface IInterviewFilterCriteria {
  candidateId?: number;
  jobRequisitionId?: number;
  interviewType?: InterviewType[];
  status?: InterviewStatus[];
  interviewerId?: number;
  fromDate?: Date;
  toDate?: Date;
  result?: InterviewResult[];
}

// ==================== Dashboard Statistics ====================

export interface ITalentDashboardStats {
  activeCandidates: number;
  openPositions: number;
  interviewsThisWeek: number;
  pendingOffers: number;
  avgTimeToHire: number;
  offerAcceptanceRate: number;
}

// ==================== Bulk Operations ====================

export interface ICandidateBulkOperation {
  operation: 'ChangeStatus' | 'AssignRecruiter' | 'AddToPool' | 'SendEmail' | 'Archive';
  candidateIds: number[];
  parameters?: {
    status?: CandidateStatus;
    recruiterId?: number;
    poolId?: number;
    emailTemplateId?: number;
  };
}

// ==================== Interview Schedule Conflict ====================

export interface IScheduleConflict {
  interviewerId: number;
  conflictingInterviews: IInterview[];
  suggestedTimes: Date[];
}

// ==================== Hiring Pipeline Stage ====================

export interface IPipelineStage {
  status: CandidateStatus;
  count: number;
  avgDaysInStage: number;
  conversionRate: number; // To next stage
}

// ==================== Candidate Score Card ====================

export interface ICandidateScoreCard {
  candidateId: number;
  candidateName: string;
  overallScore: number;
  interviewScores: {
    interviewId: number;
    interviewType: InterviewType;
    score: number;
    interviewer: string;
    date: Date;
  }[];
  technicalAverage: number;
  behavioralAverage: number;
  recommendation: 'Strong Hire' | 'Hire' | 'Maybe' | 'No Hire';
  hiringManagerRecommendation?: string;
}
