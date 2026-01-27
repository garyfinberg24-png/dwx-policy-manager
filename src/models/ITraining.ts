// Training & Skills Builder Data Models
// Complete type definitions for the JML Training & Skills Builder module

import { IBaseListItem } from './ICommon';

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Course delivery types
 */
export enum CourseType {
  eLearning = 'eLearning',
  ILT = 'Instructor-Led',
  VirtualILT = 'Virtual Instructor-Led',
  Blended = 'Blended',
  External = 'External',
  OnTheJob = 'On-the-Job',
  Video = 'Video',
  Document = 'Document'
}

/**
 * Content format types
 */
export enum ContentFormat {
  SCORM = 'SCORM',
  xAPI = 'xAPI',
  Video = 'Video',
  Document = 'Document',
  URL = 'URL',
  Interactive = 'Interactive',
  Assessment = 'Assessment',
  Presentation = 'Presentation'
}

/**
 * Difficulty levels for courses
 */
export enum DifficultyLevel {
  Beginner = 'Beginner',
  Intermediate = 'Intermediate',
  Advanced = 'Advanced',
  Expert = 'Expert'
}

/**
 * Training enrollment status
 */
export enum EnrollmentStatus {
  NotStarted = 'Not Started',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Failed = 'Failed',
  Expired = 'Expired',
  Waived = 'Waived',
  OnHold = 'On Hold'
}

/**
 * How training was assigned
 */
export enum AssignmentType {
  Mandatory = 'Mandatory',
  Recommended = 'Recommended',
  SelfEnrolled = 'Self-Enrolled',
  ManagerAssigned = 'Manager-Assigned',
  ProcessTriggered = 'Process-Triggered',
  RoleBased = 'Role-Based'
}

/**
 * Learning path types
 */
export enum LearningPathType {
  Onboarding = 'Onboarding',
  RoleBased = 'Role-Based',
  Certification = 'Certification Prep',
  Development = 'Professional Development',
  Compliance = 'Compliance',
  Leadership = 'Leadership',
  Technical = 'Technical',
  Custom = 'Custom'
}

/**
 * Certification types
 */
export enum CertificationType {
  Industry = 'Industry',
  Vendor = 'Vendor',
  Internal = 'Internal',
  License = 'Professional License',
  Accreditation = 'Accreditation',
  Regulatory = 'Regulatory'
}

/**
 * Certification status
 */
export enum CertificationStatus {
  Active = 'Active',
  Expired = 'Expired',
  Revoked = 'Revoked',
  Pending = 'Pending Verification',
  Suspended = 'Suspended'
}

/**
 * Skill proficiency levels (1-5 scale)
 */
export enum ProficiencyLevel {
  Novice = 1,
  Beginner = 2,
  Competent = 3,
  Proficient = 4,
  Expert = 5
}

/**
 * Skill domains
 */
export enum SkillDomain {
  Technical = 'Technical',
  Leadership = 'Leadership',
  Communication = 'Communication',
  Business = 'Business',
  Creative = 'Creative',
  Analytical = 'Analytical',
  Interpersonal = 'Interpersonal',
  Industry = 'Industry-Specific'
}

/**
 * Training request status
 */
export enum TrainingRequestStatus {
  Draft = 'Draft',
  Submitted = 'Submitted',
  PendingManagerApproval = 'Pending Manager Approval',
  PendingHRApproval = 'Pending HR Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Completed = 'Completed',
  Cancelled = 'Cancelled'
}

/**
 * Training session status
 */
export enum SessionStatus {
  Scheduled = 'Scheduled',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Postponed = 'Postponed'
}

/**
 * Achievement categories
 */
export enum AchievementCategory {
  Completion = 'Completion',
  Streak = 'Streak',
  Score = 'Score',
  Social = 'Social',
  Milestone = 'Milestone',
  Special = 'Special'
}

/**
 * Skill assessment source
 */
export enum SkillSource {
  Self = 'Self-Assessment',
  Manager = 'Manager Assessment',
  Training = 'Training Completion',
  Certification = 'Certification',
  Project = 'Project Experience',
  Import = 'External Import',
  AI = 'AI Inferred'
}

// ============================================================================
// TRAINING CATALOG INTERFACES
// ============================================================================

/**
 * Training Category
 */
export interface ITrainingCategory extends IBaseListItem {
  ParentCategoryId?: number;
  ParentCategory?: ITrainingCategory;
  Description?: string;
  Icon?: string;
  Color?: string;
  SortOrder: number;
  IsActive: boolean;
  CourseCount?: number;
  Children?: ITrainingCategory[];
}

/**
 * Skill definition
 */
export interface ISkill extends IBaseListItem {
  SkillCode: string;
  Description: string;
  Domain: SkillDomain;
  CategoryId?: number;
  Category?: ITrainingCategory;
  ProficiencyLevels: IProficiencyLevelDefinition[];
  IndustryMapping?: string;
  RelatedSkillIds?: number[];
  RelatedSkills?: ISkill[];
  IsCore: boolean;
  IsActive: boolean;
  UsageCount?: number;
}

/**
 * Definition of what each proficiency level means for a skill
 */
export interface IProficiencyLevelDefinition {
  level: ProficiencyLevel;
  title: string;
  description: string;
  behavioralIndicators: string[];
}

/**
 * Training Course
 */
export interface ITrainingCourse extends IBaseListItem {
  Description: string;
  CourseCode?: string;
  CourseType: CourseType;
  ContentFormat: ContentFormat;
  ContentUrl?: string;
  Duration: number;
  DifficultyLevel: DifficultyLevel;
  CategoryId?: number;
  Category?: ITrainingCategory;
  SubCategory?: string;
  SkillIds?: number[];
  Skills?: ISkill[];
  PrerequisiteIds?: number[];
  Prerequisites?: ITrainingCourse[];
  Provider?: string;
  Language: string;
  Thumbnail?: string;
  Version: string;
  IsActive: boolean;
  IsMandatory: boolean;
  CertificationValue?: number;
  PassingScore?: number;
  MaxAttempts?: number;
  ValidityPeriod?: number;
  Tags?: string[];
  EstimatedCost?: number;
  AverageRating?: number;
  TotalEnrollments?: number;
  CompletionRate?: number;
  LearningObjectives?: string[];
  TargetAudience?: string;
  LastUpdatedDate?: Date;
}

/**
 * Course form for creating/editing
 */
export interface ITrainingCourseForm {
  Title: string;
  Description: string;
  CourseCode?: string;
  CourseType: CourseType;
  ContentFormat: ContentFormat;
  ContentUrl?: string;
  Duration: number;
  DifficultyLevel: DifficultyLevel;
  CategoryId?: number;
  SubCategory?: string;
  SkillIds?: number[];
  PrerequisiteIds?: number[];
  Provider?: string;
  Language: string;
  Thumbnail?: string;
  Version: string;
  IsActive: boolean;
  IsMandatory: boolean;
  CertificationValue?: number;
  PassingScore?: number;
  MaxAttempts?: number;
  ValidityPeriod?: number;
  Tags?: string[];
  EstimatedCost?: number;
  LearningObjectives?: string[];
  TargetAudience?: string;
}

// ============================================================================
// LEARNING PATH INTERFACES
// ============================================================================

/**
 * Course entry within a learning path
 */
export interface ILearningPathCourse {
  courseId: number;
  courseName?: string;
  order: number;
  isRequired: boolean;
  unlockConditions?: IUnlockCondition[];
  estimatedDuration?: number;
}

/**
 * Conditions for unlocking content
 */
export interface IUnlockCondition {
  type: 'course_complete' | 'assessment_pass' | 'date_reached' | 'approval' | 'milestone_reached';
  value: string | number;
  courseId?: number;
  minimumScore?: number;
}

/**
 * Milestone within a learning path
 */
export interface ILearningPathMilestone {
  id: string;
  title: string;
  description: string;
  coursesRequired: number[];
  badgeIcon?: string;
  badgeName?: string;
  celebrationMessage?: string;
  pointsAwarded?: number;
}

/**
 * Criteria for path completion
 */
export interface ICompletionCriteria {
  requiredCourses: number[];
  optionalMinimum?: number;
  assessmentRequired?: boolean;
  minimumScore?: number;
  minimumTimeSpent?: number;
}

/**
 * Certification definition
 */
export interface ICertification extends IBaseListItem {
  Provider: string;
  CertificationType: CertificationType;
  Description: string;
  ValidityPeriod?: number;
  RenewalRequirements?: string;
  PreparatoryTrainingIds?: number[];
  PreparatoryTraining?: ITrainingCourse[];
  SkillIds?: number[];
  Skills?: ISkill[];
  CostEstimate?: number;
  Url?: string;
  LogoUrl?: string;
  IsActive: boolean;
  HolderCount?: number;
  DifficultyLevel?: DifficultyLevel;
  ExamFormat?: string;
}

/**
 * Learning Path
 */
export interface ILearningPath extends IBaseListItem {
  Description: string;
  PathType: LearningPathType;
  TargetRoles?: string[];
  TargetDepartments?: string[];
  Thumbnail?: string;
  EstimatedDuration: number;
  Courses: ILearningPathCourse[];
  Milestones?: ILearningPathMilestone[];
  CompletionCriteria?: ICompletionCriteria;
  IsActive: boolean;
  IsAutoAssigned: boolean;
  Version: number;
  TotalEnrollments?: number;
  AverageCompletionTime?: number;
  CompletionRate?: number;
  SkillsGained?: ISkill[];
  CertificationsEarned?: ICertification[];
}

/**
 * Learning path form
 */
export interface ILearningPathForm {
  Title: string;
  Description: string;
  PathType: LearningPathType;
  TargetRoles?: string[];
  TargetDepartments?: string[];
  Thumbnail?: string;
  Courses: ILearningPathCourse[];
  Milestones?: ILearningPathMilestone[];
  CompletionCriteria?: ICompletionCriteria;
  IsActive: boolean;
  IsAutoAssigned: boolean;
}

// ============================================================================
// ENROLLMENT INTERFACES
// ============================================================================

/**
 * Training Enrollment
 */
export interface ITrainingEnrollment extends IBaseListItem {
  UserId: number;
  UserEmail: string;
  UserDisplayName?: string;
  UserDepartment?: string;
  CourseId: number;
  Course?: ITrainingCourse;
  LearningPathId?: number;
  LearningPath?: ILearningPath;
  ProcessId?: number;
  TaskAssignmentId?: number;
  AssignmentType: AssignmentType;
  AssignedById?: number;
  AssignedByName?: string;
  AssignedDate: Date;
  DueDate?: Date;
  Status: EnrollmentStatus;
  Progress: number;
  StartedDate?: Date;
  CompletedDate?: Date;
  Score?: number;
  AttemptCount: number;
  TimeSpent: number;
  LastAccessDate?: Date;
  CertificateUrl?: string;
  Notes?: string;
  WaivedById?: number;
  WaivedByName?: string;
  WaivedReason?: string;
  ExpirationDate?: Date;
  RemindersSent: number;
  LastReminderDate?: Date;
  IsOverdue?: boolean;
  DaysUntilDue?: number;
  DaysOverdue?: number;
}

/**
 * Enrollment options when enrolling a user
 */
export interface IEnrollmentOptions {
  assignmentType: AssignmentType;
  dueDate?: Date;
  assignedById?: number;
  processId?: number;
  taskAssignmentId?: number;
  learningPathId?: number;
  notes?: string;
}

/**
 * Progress for a single course within a path
 */
export interface ICourseProgress {
  courseId: number;
  courseName: string;
  order: number;
  isRequired: boolean;
  isUnlocked: boolean;
  status: EnrollmentStatus;
  progress: number;
  score?: number;
  completedDate?: Date;
}

/**
 * User's progress through a learning path
 */
export interface ILearnerPathProgress {
  pathId: number;
  path: ILearningPath;
  userId: number;
  enrollmentDate: Date;
  progress: number;
  coursesCompleted: number;
  coursesTotal: number;
  courseProgress: ICourseProgress[];
  milestonesReached: string[];
  currentCourse?: ITrainingCourse;
  estimatedCompletion?: Date;
  timeSpent: number;
  status: EnrollmentStatus;
}

// ============================================================================
// USER SKILLS & CERTIFICATIONS
// ============================================================================

/**
 * Evidence supporting a skill claim
 */
export interface ISkillEvidence {
  type: 'course' | 'certification' | 'project' | 'document' | 'external';
  id?: number;
  title: string;
  date: Date;
  url?: string;
  description?: string;
}

/**
 * Peer endorsement for a skill
 */
export interface ISkillEndorsement {
  endorserId: number;
  endorserName: string;
  endorserEmail?: string;
  endorserTitle?: string;
  date: Date;
  comment?: string;
  level?: ProficiencyLevel;
}

/**
 * User's skill record
 */
export interface IUserSkill extends IBaseListItem {
  UserId: number;
  UserEmail: string;
  SkillId: number;
  Skill?: ISkill;
  SelfRating?: ProficiencyLevel;
  ManagerRating?: ProficiencyLevel;
  VerifiedLevel?: ProficiencyLevel;
  LastAssessedDate?: Date;
  AssessedById?: number;
  AssessedByName?: string;
  Evidence?: ISkillEvidence[];
  Endorsements?: ISkillEndorsement[];
  Source: SkillSource;
  Notes?: string;
  EffectiveLevel?: ProficiencyLevel;
  GapFromRequired?: number;
}

/**
 * User's certification record
 */
export interface IUserCertification extends IBaseListItem {
  UserId: number;
  UserEmail: string;
  UserDisplayName?: string;
  CertificationId: number;
  Certification?: ICertification;
  CredentialNumber?: string;
  IssueDate: Date;
  ExpirationDate?: Date;
  Status: CertificationStatus;
  VerificationUrl?: string;
  CertificateFile?: string;
  RenewalDate?: Date;
  ApprovalStatus?: 'Pending' | 'Approved' | 'Rejected';
  ApprovedById?: number;
  ApprovedByName?: string;
  Notes?: string;
  RemindersSent: number;
  DaysUntilExpiration?: number;
  IsExpiringSoon?: boolean;
  IsExpired?: boolean;
}

/**
 * Form for adding user certification
 */
export interface IUserCertificationForm {
  UserId: number;
  CertificationId: number;
  CredentialNumber?: string;
  IssueDate: Date;
  ExpirationDate?: Date;
  VerificationUrl?: string;
  CertificateFile?: string;
  Notes?: string;
}

// ============================================================================
// ROLE COMPETENCIES
// ============================================================================

/**
 * Skill requirement for a role
 */
export interface IRoleSkillRequirement {
  skillId: number;
  skillName: string;
  skillCode?: string;
  requiredLevel: ProficiencyLevel;
  weight: 'Critical' | 'Important' | 'Preferred';
}

/**
 * Certification requirement for a role
 */
export interface IRoleCertificationRequirement {
  certificationId: number;
  certificationName: string;
  isRequired: boolean;
  alternateIds?: number[];
}

/**
 * Step in a career ladder
 */
export interface ICareerStep {
  roleId?: number;
  roleName: string;
  level: string;
  order: number;
  typicalTimeframe?: string;
  keySkillsNeeded?: string[];
}

/**
 * Role competency definition
 */
export interface IRoleCompetency extends IBaseListItem {
  RoleTitle: string;
  RoleFamily?: string;
  Department?: string;
  Level: string;
  RequiredSkills: IRoleSkillRequirement[];
  PreferredSkills?: IRoleSkillRequirement[];
  RequiredCertifications?: IRoleCertificationRequirement[];
  LearningPathIds?: number[];
  LearningPaths?: ILearningPath[];
  SuccessionPath?: ICareerStep[];
  IsActive: boolean;
  EffectiveDate: Date;
  Version: number;
}

// ============================================================================
// SKILLS GAP ANALYSIS
// ============================================================================

/**
 * Individual skill gap
 */
export interface ISkillGap {
  skill: ISkill;
  skillId: number;
  skillName: string;
  requiredLevel: ProficiencyLevel;
  currentLevel: ProficiencyLevel;
  gap: number;
  priority: 'Critical' | 'High' | 'Medium' | 'Low';
  suggestedCourses: ITrainingCourse[];
  estimatedTimeToClose: number;
}

/**
 * Skill where user exceeds requirements
 */
export interface ISkillStrength {
  skill: ISkill;
  skillId: number;
  skillName: string;
  currentLevel: ProficiencyLevel;
  requiredLevel?: ProficiencyLevel;
  exceeds: number;
}

/**
 * AI-generated training recommendation
 */
export interface ITrainingRecommendation {
  course: ITrainingCourse;
  reason: string;
  priority: number;
  skillsAddressed: ISkill[];
  estimatedImpact: number;
  matchScore: number;
}

/**
 * Skills gap analysis result
 */
export interface ISkillsGapAnalysis {
  userId: number;
  userName?: string;
  roleId: number;
  roleName: string;
  analysisDate: Date;
  overallReadiness: number;
  gaps: ISkillGap[];
  strengths: ISkillStrength[];
  recommendedTraining: ITrainingRecommendation[];
  recommendedCertifications: ICertification[];
  estimatedTimeToClose: number;
  priorityActions: string[];
}

/**
 * Team member's skills for matrix
 */
export interface ITeamMemberSkills {
  userId: number;
  userName: string;
  role: string;
  skills: Map<number, ProficiencyLevel>;
  overallReadiness: number;
  gapCount: number;
}

/**
 * Aggregate skill gap across team
 */
export interface IAggregateSkillGap {
  skill: ISkill;
  averageLevel: number;
  requiredLevel: ProficiencyLevel;
  gap: number;
  affectedMembers: number;
  totalMembers: number;
  priority: 'Critical' | 'High' | 'Medium' | 'Low';
}

/**
 * Team skills matrix
 */
export interface ITeamSkillsMatrix {
  managerId: number;
  teamMembers: ITeamMemberSkills[];
  skillColumns: ISkill[];
  aggregateGaps: IAggregateSkillGap[];
  teamStrengths: ISkill[];
  teamWeaknesses: ISkill[];
  analysisDate: Date;
}

// ============================================================================
// TRAINING REQUESTS & SESSIONS
// ============================================================================

/**
 * Training request
 */
export interface ITrainingRequest extends IBaseListItem {
  RequesterId: number;
  RequesterEmail: string;
  RequesterName?: string;
  RequesterDepartment?: string;
  RequestType: 'Course' | 'Certification' | 'Conference' | 'External' | 'Other';
  ItemTitle: string;
  ItemId?: number;
  ExternalUrl?: string;
  Description: string;
  BusinessJustification: string;
  EstimatedCost?: number;
  RequestedStartDate?: Date;
  RequestedEndDate?: Date;
  Status: TrainingRequestStatus;
  ManagerId?: number;
  ManagerName?: string;
  ManagerApprovalDate?: Date;
  ManagerComments?: string;
  HRApprovalRequired: boolean;
  HRApproverId?: number;
  HRApproverName?: string;
  HRApprovalDate?: Date;
  HRComments?: string;
  RejectionReason?: string;
  CompletedDate?: Date;
  ActualCost?: number;
  Attachments?: string[];
}

/**
 * Training request form
 */
export interface ITrainingRequestForm {
  RequestType: 'Course' | 'Certification' | 'Conference' | 'External' | 'Other';
  ItemTitle: string;
  ItemId?: number;
  ExternalUrl?: string;
  Description: string;
  BusinessJustification: string;
  EstimatedCost?: number;
  RequestedStartDate?: Date;
  RequestedEndDate?: Date;
}

/**
 * Material for a training session
 */
export interface ISessionMaterial {
  title: string;
  type: 'Presentation' | 'Document' | 'Video' | 'Link' | 'Exercise';
  url: string;
  isPreWork?: boolean;
}

/**
 * ILT/VILT Training Session
 */
export interface ITrainingSession extends IBaseListItem {
  CourseId: number;
  Course?: ITrainingCourse;
  SessionCode?: string;
  SessionType: 'InPerson' | 'Virtual' | 'Hybrid';
  InstructorId?: number;
  InstructorName?: string;
  ExternalInstructor?: string;
  StartDateTime: Date;
  EndDateTime: Date;
  TimeZone: string;
  Location?: string;
  RoomNumber?: string;
  VirtualMeetingUrl?: string;
  MaxCapacity: number;
  MinCapacity?: number;
  CurrentEnrollment: number;
  WaitlistCount: number;
  Status: SessionStatus;
  Materials?: ISessionMaterial[];
  RecordingUrl?: string;
  Notes?: string;
  CancellationDeadline?: Date;
  AvailableSeats?: number;
  IsFullyBooked?: boolean;
  CanEnroll?: boolean;
}

/**
 * Session attendee
 */
export interface ISessionAttendee {
  sessionId: number;
  userId: number;
  userName: string;
  userEmail: string;
  registrationDate: Date;
  status: 'Registered' | 'Waitlisted' | 'Attended' | 'NoShow' | 'Cancelled';
  attendanceMarked?: boolean;
  notes?: string;
}

// ============================================================================
// FEEDBACK & ASSESSMENTS
// ============================================================================

/**
 * Training feedback/rating
 */
export interface ITrainingFeedback extends IBaseListItem {
  EnrollmentId: number;
  CourseId: number;
  CourseName?: string;
  UserId: number;
  OverallRating: number;
  ContentRating?: number;
  InstructorRating?: number;
  RelevanceRating?: number;
  DifficultyRating?: number;
  Comments?: string;
  Improvements?: string;
  WouldRecommend: boolean;
  SubmittedDate: Date;
  IsAnonymous: boolean;
}

/**
 * Question option for multiple choice
 */
export interface IQuestionOption {
  id: string;
  text: string;
  isCorrect?: boolean;
}

/**
 * Assessment question
 */
export interface IAssessmentQuestion {
  id: string;
  questionType: 'MultipleChoice' | 'TrueFalse' | 'MultiSelect' | 'ShortAnswer' | 'Matching' | 'Ordering';
  questionText: string;
  options?: IQuestionOption[];
  correctAnswer?: string | string[];
  points: number;
  explanation?: string;
  skillId?: number;
  difficultyLevel?: DifficultyLevel;
}

/**
 * Assessment/Quiz definition
 */
export interface IAssessment extends IBaseListItem {
  CourseId?: number;
  AssessmentType: 'Quiz' | 'Exam' | 'Survey' | 'SelfAssessment';
  Instructions?: string;
  TimeLimit?: number;
  PassingScore: number;
  MaxAttempts: number;
  RandomizeQuestions: boolean;
  RandomizeAnswers: boolean;
  ShowResults: boolean;
  ShowCorrectAnswers: boolean;
  Questions: IAssessmentQuestion[];
  IsActive: boolean;
}

/**
 * Answer to a question
 */
export interface IQuestionAnswer {
  questionId: string;
  answer: string | string[];
  isCorrect: boolean;
  pointsEarned: number;
}

/**
 * User's assessment attempt
 */
export interface IAssessmentAttempt extends IBaseListItem {
  AssessmentId: number;
  UserId: number;
  EnrollmentId?: number;
  StartTime: Date;
  EndTime?: Date;
  Score: number;
  MaxScore: number;
  Percentage: number;
  Passed: boolean;
  Answers: IQuestionAnswer[];
  TimeSpent: number;
  AttemptNumber: number;
}

// ============================================================================
// GAMIFICATION
// ============================================================================

/**
 * Criteria for earning an achievement
 */
export interface IAchievementCriteria {
  type: 'courses_completed' | 'path_completed' | 'streak_days' | 'score_achieved' | 'certifications' | 'hours_learned' | 'skills_verified' | 'custom';
  threshold: number;
  courseIds?: number[];
  pathId?: number;
  skillIds?: number[];
  customCondition?: string;
}

/**
 * Achievement/Badge definition
 */
export interface IAchievement extends IBaseListItem {
  AchievementCode: string;
  Description: string;
  Category: AchievementCategory;
  Icon: string;
  BadgeColor?: string;
  PointsValue: number;
  Criteria: IAchievementCriteria;
  IsActive: boolean;
  IsHidden: boolean;
  Rarity?: 'Common' | 'Uncommon' | 'Rare' | 'Epic' | 'Legendary';
}

/**
 * User's earned achievement
 */
export interface IUserAchievement extends IBaseListItem {
  UserId: number;
  UserEmail: string;
  AchievementId: number;
  Achievement?: IAchievement;
  EarnedDate: Date;
  Progress?: number;
  IsDisplayed: boolean;
}

/**
 * Leaderboard entry
 */
export interface ILeaderboardEntry {
  rank: number;
  userId: number;
  userName: string;
  userDepartment?: string;
  points: number;
  coursesCompleted: number;
  achievementsEarned: number;
  currentStreak: number;
  avatarUrl?: string;
}

/**
 * User's gamification stats
 */
export interface IUserGamificationStats {
  userId: number;
  totalPoints: number;
  rank?: number;
  currentStreak: number;
  longestStreak: number;
  lastActivityDate?: Date;
  coursesCompleted: number;
  certificationsEarned: number;
  achievementsEarned: number;
  hoursLearned: number;
  level?: number;
  levelProgress?: number;
}

// ============================================================================
// DASHBOARD & ANALYTICS
// ============================================================================

/**
 * Learner dashboard data
 */
export interface ILearnerDashboard {
  user: {
    id: number;
    name: string;
    email: string;
    role: string;
    department: string;
    manager?: string;
    avatarUrl?: string;
  };
  currentEnrollments: ITrainingEnrollment[];
  completedRecently: ITrainingEnrollment[];
  upcomingDue: ITrainingEnrollment[];
  overdue: ITrainingEnrollment[];
  learningPaths: ILearnerPathProgress[];
  skills: IUserSkill[];
  certifications: IUserCertification[];
  achievements: IUserAchievement[];
  gamificationStats: IUserGamificationStats;
  recommendations: ITrainingRecommendation[];
  upcomingSessions: ITrainingSession[];
  totalTrainingHours: number;
  coursesCompletedThisYear: number;
  averageScore: number;
  complianceStatus: number;
}

/**
 * Team member summary for manager view
 */
export interface ITeamMemberSummary {
  userId: number;
  userName: string;
  role: string;
  complianceRate: number;
  inProgressCount: number;
  overdueCount: number;
  lastActivity?: Date;
  skillReadiness: number;
  avatarUrl?: string;
}

/**
 * Course popularity metric
 */
export interface ICoursePopularity {
  courseId: number;
  courseName: string;
  enrollments: number;
  completions: number;
  averageRating: number;
  averageScore: number;
}

/**
 * Trend data point for charts
 */
export interface ITrendDataPoint {
  date: Date;
  label: string;
  value: number;
  secondaryValue?: number;
}

/**
 * Manager team dashboard data
 */
export interface IManagerDashboard {
  managerId: number;
  managerName: string;
  teamSize: number;
  teamMembers: ITeamMemberSummary[];
  teamComplianceRate: number;
  totalOverdue: number;
  totalInProgress: number;
  pendingApprovals: ITrainingRequest[];
  teamSkillsOverview: ITeamSkillsMatrix;
  completionTrend: ITrendDataPoint[];
  topCourses: ICoursePopularity[];
  upcomingExpiryCertifications: IUserCertification[];
}

/**
 * Department training metric
 */
export interface IDepartmentMetric {
  department: string;
  totalEmployees: number;
  activeEnrollments: number;
  completionRate: number;
  complianceRate: number;
  averageSkillLevel: number;
  trainingHours: number;
}

/**
 * Enrollment by type metric
 */
export interface IEnrollmentTypeMetric {
  type: AssignmentType;
  count: number;
  percentage: number;
  completionRate: number;
}

/**
 * Admin dashboard metrics
 */
export interface IAdminDashboardMetrics {
  totalCourses: number;
  activeCourses: number;
  totalEnrollments: number;
  activeEnrollments: number;
  completedThisMonth: number;
  averageCompletionRate: number;
  overdueCount: number;
  complianceRate: number;
  totalTrainingHours: number;
  averageRating: number;
  certificationExpiringCount: number;
  pendingRequests: number;
  topCourses: ICoursePopularity[];
  completionTrend: ITrendDataPoint[];
  departmentMetrics: IDepartmentMetric[];
  enrollmentsByType: IEnrollmentTypeMetric[];
}

// ============================================================================
// FILTERS & SEARCH
// ============================================================================

/**
 * Course catalog filters
 */
export interface ICourseFilters {
  searchQuery?: string;
  categoryIds?: number[];
  courseTypes?: CourseType[];
  difficultyLevels?: DifficultyLevel[];
  providers?: string[];
  languages?: string[];
  minDuration?: number;
  maxDuration?: number;
  skillIds?: number[];
  isMandatory?: boolean;
  hasAvailableSeats?: boolean;
  tags?: string[];
  sortBy?: 'title' | 'rating' | 'popularity' | 'duration' | 'date';
  sortOrder?: 'asc' | 'desc';
}

/**
 * Enrollment filters
 */
export interface IEnrollmentFilters {
  userId?: number;
  courseId?: number;
  pathId?: number;
  status?: EnrollmentStatus[];
  assignmentType?: AssignmentType[];
  isOverdue?: boolean;
  dueDateFrom?: Date;
  dueDateTo?: Date;
  completedDateFrom?: Date;
  completedDateTo?: Date;
  departmentIds?: string[];
  managerId?: number;
}

/**
 * Skills filters
 */
export interface ISkillsFilters {
  domains?: SkillDomain[];
  categoryIds?: number[];
  isCore?: boolean;
  searchQuery?: string;
  hasGap?: boolean;
  minLevel?: ProficiencyLevel;
  maxLevel?: ProficiencyLevel;
}

// ============================================================================
// JML INTEGRATION
// ============================================================================

/**
 * Training integration with JML Process
 */
export interface IProcessTrainingConfig {
  processType: 'Joiner' | 'Mover' | 'Leaver';
  roleId?: number;
  departmentId?: string;
  learningPathIds: number[];
  mandatoryCourseIds: number[];
  recommendedCourseIds: number[];
  certificationIds: number[];
  skillAssessmentRequired: boolean;
  daysToComplete: number;
}

/**
 * Training task in JML process
 */
export interface ITrainingTaskMapping {
  taskAssignmentId: number;
  enrollmentId: number;
  courseId: number;
  status: EnrollmentStatus;
  progress: number;
  syncedDate: Date;
}

// ============================================================================
// EXPORT & NOTIFICATIONS
// ============================================================================

/**
 * Training report export options
 */
export interface ITrainingExportOptions {
  format: 'PDF' | 'Excel' | 'CSV';
  reportType: 'enrollments' | 'completions' | 'compliance' | 'skills' | 'certifications';
  dateRange?: { from: Date; to: Date };
  filters?: IEnrollmentFilters | ISkillsFilters;
  includeCharts?: boolean;
  groupBy?: string;
}

/**
 * Certificate export data
 */
export interface ICertificateExport {
  enrollmentId: number;
  userId: number;
  userName: string;
  courseName: string;
  completionDate: Date;
  score?: number;
  certificateNumber: string;
  issueDate: Date;
  validUntil?: Date;
}

/**
 * Training notification configuration
 */
export interface ITrainingNotificationConfig {
  dueDateReminders: number[];
  expirationReminders: number[];
  enableOverdueAlerts: boolean;
  overdueAlertFrequency: number;
  enableCompletionNotifications: boolean;
  enableManagerDigest: boolean;
  managerDigestFrequency: 'daily' | 'weekly';
  enableAchievementNotifications: boolean;
}

/**
 * Training notification
 */
export interface ITrainingNotification {
  id: string;
  type: 'due_reminder' | 'overdue_alert' | 'completion' | 'expiration' | 'achievement' | 'assignment' | 'approval';
  userId: number;
  title: string;
  message: string;
  link?: string;
  enrollmentId?: number;
  certificationId?: number;
  isRead: boolean;
  createdDate: Date;
}
