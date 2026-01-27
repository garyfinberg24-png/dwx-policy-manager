// Onboarding Models
// Interfaces for user onboarding, tutorials, and help system

import { IBaseListItem } from './ICommon';

/**
 * Onboarding Tutorial
 */
export interface IOnboardingTutorial extends IBaseListItem {
  TutorialType: TutorialType;
  TargetAudience: string[]; // Roles or departments
  Priority: number; // Display order
  IsActive: boolean;
  IsMandatory?: boolean;
  Steps: ITutorialStep[];
  VideoUrl?: string;
  Duration?: number; // Estimated duration in minutes
  CompletionCriteria?: string; // JSON criteria
  Tags?: string[];
}

/**
 * Tutorial Types
 */
export enum TutorialType {
  GettingStarted = 'Getting Started',
  CreateProcess = 'Create Process',
  ManageTasks = 'Manage Tasks',
  Approvals = 'Approvals',
  Templates = 'Templates',
  Reporting = 'Reporting',
  Integrations = 'Integrations',
  AIFeatures = 'AI Features',
  Advanced = 'Advanced Features'
}

/**
 * Tutorial Step
 */
export interface ITutorialStep {
  id: string;
  title: string;
  description: string;
  content?: string; // Markdown or HTML
  targetElement?: string; // CSS selector for overlay
  position?: 'top' | 'right' | 'bottom' | 'left' | 'center';
  action?: ITutorialAction;
  videoUrl?: string;
  imageUrl?: string;
  duration?: number; // Seconds to display
  canSkip?: boolean;
  validationRequired?: boolean;
  validationCriteria?: string; // JSON
}

/**
 * Tutorial Action
 */
export interface ITutorialAction {
  type: 'click' | 'navigate' | 'input' | 'wait' | 'highlight';
  selector?: string; // CSS selector
  value?: string; // For input actions
  url?: string; // For navigate actions
  waitFor?: number; // Milliseconds for wait actions
}

/**
 * User Tutorial Progress
 */
export interface IUserTutorialProgress extends IBaseListItem {
  UserId: string;
  UserEmail: string;
  TutorialId: number;
  TutorialType: TutorialType;
  Status: TutorialStatus;
  CurrentStep: number;
  CompletedSteps: number[];
  StartedDate: Date;
  CompletedDate?: Date;
  TimeSpent?: number; // Minutes
  SkippedSteps?: number[];
  Score?: number; // If tutorial has quiz/validation
}

/**
 * Tutorial Status
 */
export enum TutorialStatus {
  NotStarted = 'Not Started',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Skipped = 'Skipped',
  Failed = 'Failed'
}

/**
 * Contextual Help
 */
export interface IContextualHelp extends IBaseListItem {
  PageUrl: string; // Pattern matching for page
  ElementSelector?: string; // CSS selector
  HelpText: string; // Short help text
  DetailedHelp?: string; // Full help content (Markdown)
  VideoUrl?: string;
  RelatedArticles?: IHelpArticle[];
  Position: 'top' | 'right' | 'bottom' | 'left';
  TriggerEvent?: 'hover' | 'focus' | 'click';
  Priority: number;
  IsActive: boolean;
  ShowCount?: number; // Max times to show (0 = always)
}

/**
 * Help Article
 */
export interface IHelpArticle extends IBaseListItem {
  Category: HelpCategory;
  Subcategory?: string;
  Content: string; // Markdown content
  VideoUrl?: string;
  Attachments?: string[]; // URLs to files
  Tags?: string[];
  ViewCount?: number;
  HelpfulVotes?: number;
  UnhelpfulVotes?: number;
  RelatedArticles?: number[]; // Article IDs
  IsPublished: boolean;
  PublishedDate?: Date;
  LastUpdated?: Date;
}

/**
 * Help Categories
 */
export enum HelpCategory {
  GettingStarted = 'Getting Started',
  Processes = 'Processes',
  Tasks = 'Tasks',
  Approvals = 'Approvals',
  Templates = 'Templates',
  Integrations = 'Integrations',
  Reporting = 'Reporting & Analytics',
  AIFeatures = 'AI Features',
  Administration = 'Administration',
  Troubleshooting = 'Troubleshooting',
  FAQ = 'FAQ'
}

/**
 * What's New Announcement
 */
export interface IWhatsNew extends IBaseListItem {
  Version: string; // e.g., "2.5.0"
  ReleaseDate: Date;
  Type: AnnouncementType;
  Priority: 'low' | 'medium' | 'high' | 'critical';
  Title: string;
  Summary: string;
  DetailedDescription?: string; // Markdown
  ImageUrl?: string;
  VideoUrl?: string;
  LearnMoreUrl?: string;
  IsActive: boolean;
  ShowUntil?: Date; // Auto-hide after date
  TargetRoles?: string[]; // Show to specific roles only
  RequiresDismissal?: boolean;
}

/**
 * Announcement Types
 */
export enum AnnouncementType {
  NewFeature = 'New Feature',
  Enhancement = 'Enhancement',
  BugFix = 'Bug Fix',
  Maintenance = 'Maintenance',
  Deprecation = 'Deprecation',
  Security = 'Security Update',
  Important = 'Important Notice'
}

/**
 * User Announcement View
 */
export interface IUserAnnouncementView extends IBaseListItem {
  UserId: string;
  AnnouncementId: number;
  ViewedDate: Date;
  Dismissed: boolean;
  DismissedDate?: Date;
  FeedbackRating?: number; // 1-5 stars
  FeedbackComment?: string;
}

/**
 * Tooltip Configuration
 */
export interface ITooltip {
  id: string;
  selector: string; // CSS selector
  content: string;
  title?: string;
  position?: 'top' | 'right' | 'bottom' | 'left' | 'auto';
  trigger?: 'hover' | 'click' | 'focus';
  maxWidth?: number;
  showArrow?: boolean;
  showIcon?: boolean;
  icon?: string;
  persistent?: boolean; // Stays open until closed
}

/**
 * Interactive Tour
 */
export interface IInteractiveTour {
  id: string;
  name: string;
  description: string;
  steps: ITourStep[];
  autoStart?: boolean;
  showProgress?: boolean;
  allowSkip?: boolean;
  completionAction?: () => void;
}

/**
 * Tour Step
 */
export interface ITourStep {
  target: string; // CSS selector
  title: string;
  content: string;
  placement?: 'top' | 'right' | 'bottom' | 'left' | 'center';
  highlightSelector?: string;
  beforeShow?: () => Promise<void>;
  afterShow?: () => void;
  beforeHide?: () => void;
  buttons?: ITourButton[];
}

/**
 * Tour Button
 */
export interface ITourButton {
  text: string;
  action: 'next' | 'back' | 'skip' | 'finish' | 'custom';
  classes?: string;
  onClick?: () => void;
}

/**
 * Sample Data Template
 */
export interface ISampleDataTemplate extends IBaseListItem {
  TemplateType: 'Process' | 'Task' | 'Approval' | 'Template' | 'Full Demo';
  Description: string;
  DataSet: string; // JSON of sample data
  IsActive: boolean;
  InstallCount?: number;
  LastUsed?: Date;
}

/**
 * Onboarding Checklist
 */
export interface IOnboardingChecklist {
  userId: string;
  items: IOnboardingChecklistItem[];
  completedCount: number;
  totalCount: number;
  percentComplete: number;
  startedDate: Date;
  completedDate?: Date;
}

/**
 * Onboarding Checklist Item
 */
export interface IOnboardingChecklistItem {
  id: string;
  title: string;
  description: string;
  category: 'setup' | 'learn' | 'practice' | 'explore';
  isCompleted: boolean;
  completedDate?: Date;
  isRequired: boolean;
  action?: {
    label: string;
    url?: string;
    handler?: () => void;
  };
  tutorialId?: number;
  estimatedTime?: number; // Minutes
}

/**
 * Guided Tour Configuration
 */
export interface IGuidedTourConfig {
  showOnFirstVisit: boolean;
  showOnFeatureRelease: boolean;
  autoStartDelay?: number; // Milliseconds
  maxShowCount?: number;
  dismissable: boolean;
  theme?: 'light' | 'dark' | 'auto';
  progressIndicator: boolean;
  keyboard: boolean; // Allow keyboard navigation
}

/**
 * User Help Interaction
 */
export interface IUserHelpInteraction extends IBaseListItem {
  UserId: string;
  InteractionType: HelpInteractionType;
  HelpArticleId?: number;
  TutorialId?: number;
  TooltipId?: string;
  SearchQuery?: string;
  WasHelpful?: boolean;
  FeedbackComment?: string;
  TimeSpent?: number; // Seconds
  PageUrl?: string;
}

/**
 * Help Interaction Types
 */
export enum HelpInteractionType {
  ArticleViewed = 'Article Viewed',
  TutorialStarted = 'Tutorial Started',
  TutorialCompleted = 'Tutorial Completed',
  TooltipViewed = 'Tooltip Viewed',
  SearchPerformed = 'Search Performed',
  VideoWatched = 'Video Watched',
  FeedbackSubmitted = 'Feedback Submitted',
  SampleDataInstalled = 'Sample Data Installed'
}

/**
 * Feature Highlight
 */
export interface IFeatureHighlight {
  id: string;
  featureName: string;
  description: string;
  selector?: string; // Element to highlight
  position?: 'top' | 'right' | 'bottom' | 'left';
  imageUrl?: string;
  videoUrl?: string;
  ctaText?: string;
  ctaAction?: () => void;
  showOnce?: boolean;
  expiryDate?: Date;
}

/**
 * Quick Start Guide
 */
export interface IQuickStartGuide {
  title: string;
  description: string;
  estimatedTime: number; // Minutes
  steps: IQuickStartStep[];
  targetAudience?: string[];
}

/**
 * Quick Start Step
 */
export interface IQuickStartStep {
  title: string;
  description: string;
  action?: {
    label: string;
    url?: string;
    handler?: () => void;
  };
  isCompleted?: boolean;
}

/**
 * Video Tutorial
 */
export interface IVideoTutorial extends IBaseListItem {
  VideoUrl: string;
  ThumbnailUrl?: string;
  Duration: number; // Seconds
  Category: HelpCategory;
  Transcript?: string;
  Subtitles?: ISubtitleTrack[];
  ViewCount?: number;
  AverageRating?: number;
  IsPublished: boolean;
}

/**
 * Subtitle Track
 */
export interface ISubtitleTrack {
  language: string;
  label: string;
  url: string; // VTT file URL
}

/**
 * Default Onboarding Checklist
 */
export const DEFAULT_ONBOARDING_CHECKLIST: IOnboardingChecklistItem[] = [
  // Setup (Required)
  {
    id: 'complete-profile',
    title: 'Complete Your Profile',
    description: 'Set your timezone, language, and notification preferences',
    category: 'setup',
    isCompleted: false,
    isRequired: true,
    action: {
      label: 'Go to Settings',
      url: '/settings'
    },
    estimatedTime: 3
  },
  {
    id: 'configure-dashboard',
    title: 'Customize Your Dashboard',
    description: 'Add widgets and arrange your dashboard layout',
    category: 'setup',
    isCompleted: false,
    isRequired: false,
    action: {
      label: 'Customize Dashboard',
      url: '/dashboard'
    },
    estimatedTime: 5
  },

  // Learn
  {
    id: 'watch-intro-video',
    title: 'Watch Introduction Video',
    description: 'Learn about the JML system and its key features',
    category: 'learn',
    isCompleted: false,
    isRequired: true,
    action: {
      label: 'Watch Video'
    },
    estimatedTime: 5
  },
  {
    id: 'take-getting-started-tour',
    title: 'Take the Getting Started Tour',
    description: 'Interactive walkthrough of the main features',
    category: 'learn',
    isCompleted: false,
    isRequired: true,
    tutorialId: 1,
    estimatedTime: 10
  },
  {
    id: 'read-user-guide',
    title: 'Read the User Guide',
    description: 'Browse the documentation to learn more',
    category: 'learn',
    isCompleted: false,
    isRequired: false,
    action: {
      label: 'Open User Guide',
      url: '/help'
    },
    estimatedTime: 15
  },

  // Practice
  {
    id: 'create-first-process',
    title: 'Create Your First Process',
    description: 'Walk through creating an onboarding process',
    category: 'practice',
    isCompleted: false,
    isRequired: true,
    tutorialId: 2,
    estimatedTime: 10
  },
  {
    id: 'assign-task',
    title: 'Assign a Task',
    description: 'Learn how to assign and manage tasks',
    category: 'practice',
    isCompleted: false,
    isRequired: false,
    tutorialId: 3,
    estimatedTime: 5
  },
  {
    id: 'submit-approval',
    title: 'Submit an Approval',
    description: 'Try the approval workflow',
    category: 'practice',
    isCompleted: false,
    isRequired: false,
    tutorialId: 4,
    estimatedTime: 5
  },

  // Explore
  {
    id: 'explore-templates',
    title: 'Explore Templates',
    description: 'Browse and try different process templates',
    category: 'explore',
    isCompleted: false,
    isRequired: false,
    action: {
      label: 'View Templates',
      url: '/templates'
    },
    estimatedTime: 5
  },
  {
    id: 'try-ai-features',
    title: 'Try AI Features',
    description: 'Experience AI-powered task recommendations and predictions',
    category: 'explore',
    isCompleted: false,
    isRequired: false,
    action: {
      label: 'Explore AI Features'
    },
    estimatedTime: 10
  },
  {
    id: 'install-sample-data',
    title: 'Install Sample Data',
    description: 'Add sample processes and templates to explore',
    category: 'explore',
    isCompleted: false,
    isRequired: false,
    action: {
      label: 'Install Samples'
    },
    estimatedTime: 2
  },
  {
    id: 'join-community',
    title: 'Join the Community',
    description: 'Connect with other users and get support',
    category: 'explore',
    isCompleted: false,
    isRequired: false,
    action: {
      label: 'Visit Community Forum'
    },
    estimatedTime: 5
  }
];
