// @ts-nocheck
/**
 * Quiz Service - Fully Featured
 * Comprehensive quiz management with advanced question types,
 * question banks, scheduling, analytics, and certificate generation
 */

import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users";
import { logger } from "./LoggingService";
import { QuizLists } from "../constants/SharePointListNames";

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Question types supported by the quiz system
 */
export enum QuestionType {
  MultipleChoice = "Multiple Choice",
  TrueFalse = "True/False",
  MultipleSelect = "Multiple Select",
  ShortAnswer = "Short Answer",
  FillInBlank = "Fill in the Blank",
  Matching = "Matching",
  Ordering = "Ordering",
  RatingScale = "Rating Scale",
  Essay = "Essay",
  ImageChoice = "Image Choice",
  Hotspot = "Hotspot"
}

/**
 * Difficulty levels
 */
export enum DifficultyLevel {
  Easy = "Easy",
  Medium = "Medium",
  Hard = "Hard",
  Expert = "Expert"
}

/**
 * Quiz status
 */
export enum QuizStatus {
  Draft = "Draft",
  Published = "Published",
  Scheduled = "Scheduled",
  Archived = "Archived"
}

/**
 * Attempt status
 */
export enum AttemptStatus {
  InProgress = "In Progress",
  Completed = "Completed",
  Abandoned = "Abandoned",
  Expired = "Expired",
  PendingReview = "Pending Review"
}

/**
 * Grading type for essay/manual questions
 */
export enum GradingType {
  Automatic = "Automatic",
  Manual = "Manual",
  Hybrid = "Hybrid"
}

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Quiz definition
 */
export interface IQuiz {
  Id: number;
  Title: string;
  PolicyId: number;
  PolicyTitle: string;
  QuizDescription: string;
  PassingScore: number;
  TimeLimit: number;
  MaxAttempts: number;
  IsActive: boolean;
  QuestionCount: number;
  AverageScore?: number;
  CompletionRate?: number;
  QuizCategory: string;
  DifficultyLevel: string;
  RandomizeQuestions: boolean;
  RandomizeOptions: boolean;
  ShowCorrectAnswers: boolean;
  ShowExplanations: boolean;
  AllowReview: boolean;
  Status: QuizStatus;
  GradingType: GradingType;

  // Scheduling
  ScheduledStartDate?: string;
  ScheduledEndDate?: string;

  // Question Bank
  QuestionBankId?: number;
  QuestionPoolSize?: number;

  // Certificate
  GenerateCertificate: boolean;
  CertificateTemplateId?: number;

  // Advanced settings
  PerQuestionTimeLimit?: number;
  AllowPartialCredit: boolean;
  ShuffleWithinSections: boolean;
  RequireSequentialCompletion: boolean;

  // Metadata
  Tags?: string;
  CreatedDate?: string;
  ModifiedDate?: string;
  CreatedById?: number;
}

/**
 * Quiz question with all types supported
 */
export interface IQuizQuestion {
  Id: number;
  QuizId: number;
  QuestionBankId?: number;

  // Core question data
  QuestionText: string;
  QuestionType: QuestionType;
  QuestionHtml?: string;

  // Options for multiple choice/select/image choice
  OptionA?: string;
  OptionB?: string;
  OptionC?: string;
  OptionD?: string;
  OptionE?: string;
  OptionF?: string;

  // For image-based questions
  OptionAImage?: string;
  OptionBImage?: string;
  OptionCImage?: string;
  OptionDImage?: string;
  QuestionImage?: string;

  // For hotspot questions
  HotspotData?: string; // JSON: { imageUrl, regions: [{x, y, width, height, isCorrect}] }

  // For matching questions
  MatchingPairs?: string; // JSON: [{left, right}]

  // For ordering questions
  OrderingItems?: string; // JSON: [{id, text, correctOrder}]

  // For fill in blank
  BlankAnswers?: string; // JSON: [{position, acceptedAnswers: []}]
  CaseSensitive?: boolean;

  // For rating scale
  ScaleMin?: number;
  ScaleMax?: number;
  ScaleLabels?: string; // JSON: {min: "label", max: "label", mid?: "label"}
  CorrectRating?: number;
  RatingTolerance?: number;

  // For essay
  MinWordCount?: number;
  MaxWordCount?: number;
  RubricId?: number;

  // Answers
  CorrectAnswer: string;
  CorrectAnswers?: string; // Semicolon-separated for multiple
  AcceptedAnswers?: string; // JSON array for flexible matching

  // Feedback
  Explanation: string;
  CorrectFeedback?: string;
  IncorrectFeedback?: string;
  PartialFeedback?: string;
  Hint?: string;

  // Scoring
  Points: number;
  PartialCreditEnabled: boolean;
  PartialCreditPercentages?: string; // JSON: {optionA: 50, optionB: 25}
  NegativeMarking: boolean;
  NegativePoints?: number;

  // Organization
  QuestionOrder: number;
  SectionId?: number;
  SectionName?: string;
  DifficultyLevel: DifficultyLevel;
  Tags?: string;
  Category?: string;

  // Time
  TimeLimit?: number; // Per-question time limit in seconds

  // Status
  IsActive: boolean;
  IsRequired: boolean;

  // Analytics
  TimesAnswered?: number;
  TimesCorrect?: number;
  AverageTime?: number;
  DiscriminationIndex?: number;
}

/**
 * Quiz attempt
 */
export interface IQuizAttempt {
  Id: number;
  QuizId: number;
  PolicyId: number;
  UserId: number;
  UserName?: string;
  UserEmail?: string;
  AttemptNumber: number;
  StartTime: string;
  EndTime?: string;
  Score: number;
  MaxScore: number;
  Percentage: number;
  Passed: boolean;
  TimeSpent?: number;
  AnswersJson?: string;
  Status: AttemptStatus;
  PointsEarned: number;

  // Review
  ReviewedById?: number;
  ReviewedDate?: string;
  ReviewNotes?: string;

  // Certificate
  CertificateGenerated: boolean;
  CertificateUrl?: string;

  // Analytics
  QuestionsAnswered: number;
  QuestionsCorrect: number;
  QuestionsPartial: number;
  QuestionsIncorrect: number;
  QuestionsSkipped: number;
}

/**
 * Individual answer in an attempt
 */
export interface IQuizAnswer {
  questionId: number;
  questionType: QuestionType;
  selectedAnswer: string;
  selectedAnswers?: string[];
  matchingAnswers?: { left: string; right: string }[];
  orderingAnswers?: string[];
  hotspotCoordinates?: { x: number; y: number };
  essayText?: string;
  ratingValue?: number;
  isCorrect: boolean;
  isPartiallyCorrect: boolean;
  pointsEarned: number;
  maxPoints: number;
  timeSpent?: number;
  feedback?: string;
  manualGrade?: number;
  manualFeedback?: string;
}

/**
 * Quiz result summary
 */
export interface IQuizResult {
  attemptId: number;
  quizId: number;
  score: number;
  maxScore: number;
  percentage: number;
  passed: boolean;
  timeSpent: number;
  answers: IQuizAnswer[];
  pointsEarned: number;
  requiresManualReview: boolean;
  pendingQuestions: number;
  certificateUrl?: string;
  feedback?: string;
  improvementAreas?: string[];
}

/**
 * Quiz statistics
 */
export interface IQuizStatistics {
  totalAttempts: number;
  uniqueUsers: number;
  passedAttempts: number;
  failedAttempts: number;
  pendingReview: number;
  averageScore: number;
  medianScore: number;
  highestScore: number;
  lowestScore: number;
  averageTimeSpent: number;
  completionRate: number;
  passRate: number;
  averageAttemptsPerUser: number;
  scoreDistribution: { range: string; count: number }[];
  questionAnalytics: IQuestionAnalytics[];
}

/**
 * Per-question analytics
 */
export interface IQuestionAnalytics {
  questionId: number;
  questionText: string;
  questionType: QuestionType;
  timesAnswered: number;
  correctRate: number;
  partialRate: number;
  incorrectRate: number;
  skippedRate: number;
  averageTime: number;
  averageScore: number;
  discriminationIndex: number;
  difficultyIndex: number;
  commonWrongAnswers: { answer: string; count: number }[];
}

/**
 * Question bank
 */
export interface IQuestionBank {
  Id: number;
  Title: string;
  Description: string;
  Category: string;
  Tags?: string;
  QuestionCount: number;
  IsPublic: boolean;
  CreatedById: number;
  CreatedDate: string;
  ModifiedDate: string;
}

/**
 * Quiz section for organizing questions
 */
export interface IQuizSection {
  Id: number;
  QuizId: number;
  Title: string;
  Description?: string;
  Order: number;
  RandomizeWithinSection: boolean;
  QuestionsRequired?: number; // If set, randomly select this many from section
}

/**
 * Essay grading rubric
 */
export interface IGradingRubric {
  Id: number;
  Title: string;
  Description: string;
  Criteria: IRubricCriterion[];
  MaxScore: number;
}

/**
 * Rubric criterion
 */
export interface IRubricCriterion {
  id: string;
  name: string;
  description: string;
  maxPoints: number;
  levels: {
    points: number;
    description: string;
  }[];
}

/**
 * Certificate template
 */
export interface ICertificateTemplate {
  Id: number;
  Title: string;
  TemplateHtml: string;
  TemplateStyles: string;
  Placeholders: string[]; // Available merge fields
  IsDefault: boolean;
}

/**
 * Generated certificate
 */
export interface ICertificate {
  Id: number;
  AttemptId: number;
  UserId: number;
  UserName: string;
  QuizTitle: string;
  Score: number;
  PassedDate: string;
  CertificateNumber: string;
  CertificateUrl: string;
  ExpiryDate?: string;
}

/**
 * Import/Export format
 */
export interface IQuizExportData {
  version: string;
  exportDate: string;
  quiz: Partial<IQuiz>;
  sections: IQuizSection[];
  questions: Partial<IQuizQuestion>[];
}

// ============================================================================
// QUIZ SERVICE
// ============================================================================

export class QuizService {
  private sp: SPFI;

  private readonly quizListName = QuizLists.POLICY_QUIZZES;
  private readonly questionListName = QuizLists.POLICY_QUIZ_QUESTIONS;
  private readonly attemptListName = QuizLists.QUIZ_ATTEMPTS;
  private readonly bankListName = QuizLists.QUESTION_BANKS;
  private readonly sectionListName = QuizLists.QUIZ_SECTIONS;
  private readonly rubricListName = QuizLists.GRADING_RUBRICS;
  private readonly certificateListName = QuizLists.QUIZ_CERTIFICATES;
  private readonly templateListName = QuizLists.CERTIFICATE_TEMPLATES;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // Quiz Management
  // ============================================================================

  /**
   * Get all quizzes with optional filtering
   */
  public async getAllQuizzes(options?: {
    status?: QuizStatus;
    category?: string;
    policyId?: number;
    includeArchived?: boolean;
  }): Promise<IQuiz[]> {
    try {
      let filter = options?.includeArchived ? "" : "IsActive eq true";

      if (options?.status) {
        filter += filter ? ` and Status eq '${options.status}'` : `Status eq '${options.status}'`;
      }
      if (options?.category) {
        filter += filter ? ` and QuizCategory eq '${options.category}'` : `QuizCategory eq '${options.category}'`;
      }
      if (options?.policyId) {
        filter += filter ? ` and PolicyId eq ${options.policyId}` : `PolicyId eq ${options.policyId}`;
      }

      let query = this.sp.web.lists.getByTitle(this.quizListName).items;
      if (filter) {
        query = query.filter(filter);
      }

      const quizzes = await query.orderBy("Title", true).top(500)();
      return quizzes as IQuiz[];
    } catch (error) {
      logger.error("QuizService", "Failed to get quizzes", error);
      return [];
    }
  }

  /**
   * Get quiz by ID with full details
   */
  public async getQuizById(quizId: number): Promise<IQuiz | null> {
    try {
      const quiz = await this.sp.web.lists
        .getByTitle(this.quizListName)
        .items.getById(quizId)();

      return quiz as IQuiz;
    } catch (error) {
      logger.error("QuizService", `Failed to get quiz ${quizId}`, error);
      return null;
    }
  }

  /**
   * Get quizzes by policy ID
   */
  public async getQuizzesByPolicy(policyId: number): Promise<IQuiz[]> {
    return this.getAllQuizzes({ policyId, status: QuizStatus.Published });
  }

  /**
   * Get scheduled quizzes that should be active now
   */
  public async getActiveScheduledQuizzes(): Promise<IQuiz[]> {
    try {
      const now = new Date().toISOString();
      const quizzes = await this.sp.web.lists
        .getByTitle(this.quizListName)
        .items.filter(
          `Status eq '${QuizStatus.Scheduled}' and ScheduledStartDate le datetime'${now}' and (ScheduledEndDate ge datetime'${now}' or ScheduledEndDate eq null)`
        )();

      return quizzes as IQuiz[];
    } catch (error) {
      logger.error("QuizService", "Failed to get scheduled quizzes", error);
      return [];
    }
  }

  /**
   * Create a new quiz
   */
  public async createQuiz(quiz: Partial<IQuiz>): Promise<IQuiz | null> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.quizListName)
        .items.add({
          Title: quiz.Title,
          PolicyId: quiz.PolicyId,
          PolicyTitle: quiz.PolicyTitle,
          QuizDescription: quiz.QuizDescription,
          PassingScore: quiz.PassingScore || 70,
          TimeLimit: quiz.TimeLimit || 30,
          MaxAttempts: quiz.MaxAttempts || 3,
          IsActive: quiz.IsActive ?? true,
          QuestionCount: 0,
          QuizCategory: quiz.QuizCategory || "General",
          DifficultyLevel: quiz.DifficultyLevel || DifficultyLevel.Medium,
          RandomizeQuestions: quiz.RandomizeQuestions ?? true,
          RandomizeOptions: quiz.RandomizeOptions ?? false,
          ShowCorrectAnswers: quiz.ShowCorrectAnswers ?? true,
          ShowExplanations: quiz.ShowExplanations ?? true,
          AllowReview: quiz.AllowReview ?? true,
          Status: quiz.Status || QuizStatus.Draft,
          GradingType: quiz.GradingType || GradingType.Automatic,
          ScheduledStartDate: quiz.ScheduledStartDate,
          ScheduledEndDate: quiz.ScheduledEndDate,
          QuestionBankId: quiz.QuestionBankId,
          QuestionPoolSize: quiz.QuestionPoolSize,
          GenerateCertificate: quiz.GenerateCertificate ?? false,
          CertificateTemplateId: quiz.CertificateTemplateId,
          PerQuestionTimeLimit: quiz.PerQuestionTimeLimit,
          AllowPartialCredit: quiz.AllowPartialCredit ?? true,
          ShuffleWithinSections: quiz.ShuffleWithinSections ?? false,
          RequireSequentialCompletion: quiz.RequireSequentialCompletion ?? false,
          Tags: quiz.Tags
        });

      logger.info("QuizService", `Quiz created: ${quiz.Title}`);
      return result.data as IQuiz;
    } catch (error) {
      logger.error("QuizService", "Failed to create quiz", error);
      return null;
    }
  }

  /**
   * Update quiz
   */
  public async updateQuiz(quizId: number, updates: Partial<IQuiz>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.quizListName)
        .items.getById(quizId)
        .update(updates);

      logger.info("QuizService", `Quiz updated: ${quizId}`);
    } catch (error) {
      logger.error("QuizService", `Failed to update quiz ${quizId}`, error);
      throw error;
    }
  }

  /**
   * Delete quiz (soft delete)
   */
  public async deleteQuiz(quizId: number): Promise<void> {
    try {
      await this.updateQuiz(quizId, {
        IsActive: false,
        Status: QuizStatus.Archived
      });
      logger.info("QuizService", `Quiz archived: ${quizId}`);
    } catch (error) {
      logger.error("QuizService", `Failed to delete quiz ${quizId}`, error);
      throw error;
    }
  }

  /**
   * Publish quiz
   */
  public async publishQuiz(quizId: number): Promise<void> {
    const quiz = await this.getQuizById(quizId);
    if (!quiz) throw new Error("Quiz not found");

    const questions = await this.getQuizQuestions(quizId);
    if (questions.length === 0) {
      throw new Error("Cannot publish quiz with no questions");
    }

    await this.updateQuiz(quizId, {
      Status: QuizStatus.Published,
      QuestionCount: questions.length
    });
  }

  /**
   * Schedule quiz for future availability
   */
  public async scheduleQuiz(
    quizId: number,
    startDate: Date,
    endDate?: Date
  ): Promise<void> {
    await this.updateQuiz(quizId, {
      Status: QuizStatus.Scheduled,
      ScheduledStartDate: startDate.toISOString(),
      ScheduledEndDate: endDate?.toISOString()
    });
  }

  // ============================================================================
  // Question Management
  // ============================================================================

  /**
   * Get all questions for a quiz
   */
  public async getQuizQuestions(
    quizId: number,
    options?: {
      randomize?: boolean;
      sectionId?: number;
      limit?: number;
    }
  ): Promise<IQuizQuestion[]> {
    try {
      let filter = `QuizId eq ${quizId} and IsActive eq true`;
      if (options?.sectionId) {
        filter += ` and SectionId eq ${options.sectionId}`;
      }

      let questions = await this.sp.web.lists
        .getByTitle(this.questionListName)
        .items.filter(filter)
        .orderBy("QuestionOrder", true)
        .top(500)() as IQuizQuestion[];

      if (options?.randomize) {
        questions = this.shuffleArray(questions);
      }

      if (options?.limit && options.limit < questions.length) {
        questions = questions.slice(0, options.limit);
      }

      return questions;
    } catch (error) {
      logger.error("QuizService", `Failed to get questions for quiz ${quizId}`, error);
      return [];
    }
  }

  /**
   * Get questions from question bank
   */
  public async getQuestionsFromBank(
    bankId: number,
    options?: {
      category?: string;
      difficulty?: DifficultyLevel;
      type?: QuestionType;
      limit?: number;
      randomize?: boolean;
    }
  ): Promise<IQuizQuestion[]> {
    try {
      let filter = `QuestionBankId eq ${bankId} and IsActive eq true`;

      if (options?.category) {
        filter += ` and Category eq '${options.category}'`;
      }
      if (options?.difficulty) {
        filter += ` and DifficultyLevel eq '${options.difficulty}'`;
      }
      if (options?.type) {
        filter += ` and QuestionType eq '${options.type}'`;
      }

      let questions = await this.sp.web.lists
        .getByTitle(this.questionListName)
        .items.filter(filter)
        .top(500)() as IQuizQuestion[];

      if (options?.randomize) {
        questions = this.shuffleArray(questions);
      }

      if (options?.limit && options.limit < questions.length) {
        questions = questions.slice(0, options.limit);
      }

      return questions;
    } catch (error) {
      logger.error("QuizService", `Failed to get questions from bank ${bankId}`, error);
      return [];
    }
  }

  /**
   * Get question by ID
   */
  public async getQuestionById(questionId: number): Promise<IQuizQuestion | null> {
    try {
      const question = await this.sp.web.lists
        .getByTitle(this.questionListName)
        .items.getById(questionId)();

      return question as IQuizQuestion;
    } catch (error) {
      logger.error("QuizService", `Failed to get question ${questionId}`, error);
      return null;
    }
  }

  /**
   * Create a new question
   */
  public async createQuestion(question: Partial<IQuizQuestion>): Promise<IQuizQuestion | null> {
    try {
      // Get current question count for ordering
      const existingQuestions = question.QuizId
        ? await this.getQuizQuestions(question.QuizId)
        : [];
      const nextOrder = existingQuestions.length + 1;

      const result = await this.sp.web.lists
        .getByTitle(this.questionListName)
        .items.add({
          Title: (question.QuestionText || "Question").substring(0, 100),
          QuizId: question.QuizId,
          QuestionBankId: question.QuestionBankId,
          QuestionText: question.QuestionText,
          QuestionType: question.QuestionType || QuestionType.MultipleChoice,
          QuestionHtml: question.QuestionHtml,
          OptionA: question.OptionA,
          OptionB: question.OptionB,
          OptionC: question.OptionC,
          OptionD: question.OptionD,
          OptionE: question.OptionE,
          OptionF: question.OptionF,
          OptionAImage: question.OptionAImage,
          OptionBImage: question.OptionBImage,
          OptionCImage: question.OptionCImage,
          OptionDImage: question.OptionDImage,
          QuestionImage: question.QuestionImage,
          HotspotData: question.HotspotData,
          MatchingPairs: question.MatchingPairs,
          OrderingItems: question.OrderingItems,
          BlankAnswers: question.BlankAnswers,
          CaseSensitive: question.CaseSensitive ?? false,
          ScaleMin: question.ScaleMin,
          ScaleMax: question.ScaleMax,
          ScaleLabels: question.ScaleLabels,
          CorrectRating: question.CorrectRating,
          RatingTolerance: question.RatingTolerance,
          MinWordCount: question.MinWordCount,
          MaxWordCount: question.MaxWordCount,
          RubricId: question.RubricId,
          CorrectAnswer: question.CorrectAnswer,
          CorrectAnswers: question.CorrectAnswers,
          AcceptedAnswers: question.AcceptedAnswers,
          Explanation: question.Explanation,
          CorrectFeedback: question.CorrectFeedback,
          IncorrectFeedback: question.IncorrectFeedback,
          PartialFeedback: question.PartialFeedback,
          Hint: question.Hint,
          Points: question.Points || 10,
          PartialCreditEnabled: question.PartialCreditEnabled ?? false,
          PartialCreditPercentages: question.PartialCreditPercentages,
          NegativeMarking: question.NegativeMarking ?? false,
          NegativePoints: question.NegativePoints,
          QuestionOrder: question.QuestionOrder || nextOrder,
          SectionId: question.SectionId,
          SectionName: question.SectionName,
          DifficultyLevel: question.DifficultyLevel || DifficultyLevel.Medium,
          Tags: question.Tags,
          Category: question.Category,
          TimeLimit: question.TimeLimit,
          IsActive: question.IsActive ?? true,
          IsRequired: question.IsRequired ?? false
        });

      // Update quiz question count
      if (question.QuizId) {
        await this.updateQuestionCount(question.QuizId);
      }

      logger.info("QuizService", `Question created for quiz ${question.QuizId}`);
      return result.data as IQuizQuestion;
    } catch (error) {
      logger.error("QuizService", "Failed to create question", error);
      return null;
    }
  }

  /**
   * Update question
   */
  public async updateQuestion(questionId: number, updates: Partial<IQuizQuestion>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.questionListName)
        .items.getById(questionId)
        .update(updates);

      logger.info("QuizService", `Question updated: ${questionId}`);
    } catch (error) {
      logger.error("QuizService", `Failed to update question ${questionId}`, error);
      throw error;
    }
  }

  /**
   * Delete question
   */
  public async deleteQuestion(questionId: number, quizId: number): Promise<void> {
    try {
      await this.updateQuestion(questionId, { IsActive: false });
      await this.updateQuestionCount(quizId);
      logger.info("QuizService", `Question deleted: ${questionId}`);
    } catch (error) {
      logger.error("QuizService", `Failed to delete question ${questionId}`, error);
      throw error;
    }
  }

  /**
   * Duplicate question
   */
  public async duplicateQuestion(questionId: number, targetQuizId?: number): Promise<IQuizQuestion | null> {
    const original = await this.getQuestionById(questionId);
    if (!original) return null;

    const duplicate: Partial<IQuizQuestion> = { ...original };
    delete duplicate.Id;
    duplicate.QuizId = targetQuizId || original.QuizId;

    return this.createQuestion(duplicate);
  }

  /**
   * Bulk create questions
   */
  public async bulkCreateQuestions(questions: Partial<IQuizQuestion>[]): Promise<number> {
    let created = 0;
    for (const question of questions) {
      const result = await this.createQuestion(question);
      if (result) created++;
    }
    return created;
  }

  /**
   * Reorder questions
   */
  public async reorderQuestions(quizId: number, questionIds: number[]): Promise<void> {
    try {
      for (let i = 0; i < questionIds.length; i++) {
        await this.updateQuestion(questionIds[i], { QuestionOrder: i + 1 });
      }
      logger.info("QuizService", `Questions reordered for quiz ${quizId}`);
    } catch (error) {
      logger.error("QuizService", "Failed to reorder questions", error);
      throw error;
    }
  }

  /**
   * Update question count for quiz
   */
  private async updateQuestionCount(quizId: number): Promise<void> {
    try {
      const questions = await this.getQuizQuestions(quizId);
      await this.updateQuiz(quizId, { QuestionCount: questions.length });
    } catch (error) {
      logger.error("QuizService", `Failed to update question count for ${quizId}`, error);
    }
  }

  // ============================================================================
  // Quiz Attempts & Grading
  // ============================================================================

  /**
   * Start a new quiz attempt
   */
  public async startQuizAttempt(
    quizId: number,
    policyId: number,
    userId: number,
    userName: string,
    userEmail: string
  ): Promise<IQuizAttempt | null> {
    try {
      // Check eligibility
      const eligibility = await this.canUserTakeQuiz(quizId, userId);
      if (!eligibility.canTake) {
        throw new Error(eligibility.reason || "Not eligible to take quiz");
      }

      const quiz = await this.getQuizById(quizId);
      if (!quiz) throw new Error("Quiz not found");

      const previousAttempts = await this.getUserQuizAttempts(quizId, userId);
      const attemptNumber = previousAttempts.length + 1;

      // Calculate max score
      const questions = await this.getQuizQuestions(quizId, {
        randomize: quiz.RandomizeQuestions,
        limit: quiz.QuestionPoolSize
      });
      const maxScore = questions.reduce((sum, q) => sum + q.Points, 0);

      const result = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.add({
          Title: `${quiz.Title} - Attempt ${attemptNumber}`,
          QuizId: quizId,
          PolicyId: policyId,
          UserId: userId,
          UserName: userName,
          UserEmail: userEmail,
          AttemptNumber: attemptNumber,
          StartTime: new Date().toISOString(),
          Score: 0,
          MaxScore: maxScore,
          Percentage: 0,
          Passed: false,
          Status: AttemptStatus.InProgress,
          PointsEarned: 0,
          CertificateGenerated: false,
          QuestionsAnswered: 0,
          QuestionsCorrect: 0,
          QuestionsPartial: 0,
          QuestionsIncorrect: 0,
          QuestionsSkipped: 0
        });

      logger.info("QuizService", `Quiz attempt started: ${quizId} by user ${userId}`);
      return result.data as IQuizAttempt;
    } catch (error) {
      logger.error("QuizService", "Failed to start quiz attempt", error);
      return null;
    }
  }

  /**
   * Submit quiz attempt and calculate results
   */
  public async submitQuizAttempt(
    attemptId: number,
    answers: IQuizAnswer[]
  ): Promise<IQuizResult | null> {
    try {
      const attempt = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)() as IQuizAttempt;

      const quiz = await this.getQuizById(attempt.QuizId);
      if (!quiz) throw new Error("Quiz not found");

      const startTime = new Date(attempt.StartTime);
      const endTime = new Date();
      const timeSpent = Math.round((endTime.getTime() - startTime.getTime()) / 1000 / 60);

      // Calculate scores
      let score = 0;
      let questionsCorrect = 0;
      let questionsPartial = 0;
      let questionsIncorrect = 0;
      let questionsSkipped = 0;
      let requiresManualReview = false;
      let pendingQuestions = 0;

      for (const answer of answers) {
        if (answer.questionType === QuestionType.Essay) {
          requiresManualReview = true;
          pendingQuestions++;
        } else {
          score += answer.pointsEarned;
          if (answer.isCorrect) {
            questionsCorrect++;
          } else if (answer.isPartiallyCorrect) {
            questionsPartial++;
          } else if (answer.selectedAnswer || answer.selectedAnswers?.length) {
            questionsIncorrect++;
          } else {
            questionsSkipped++;
          }
        }
      }

      const maxScore = attempt.MaxScore;
      const percentage = maxScore > 0 ? Math.round((score / maxScore) * 100) : 0;
      const passed = percentage >= quiz.PassingScore && !requiresManualReview;

      // Update attempt
      const status = requiresManualReview ? AttemptStatus.PendingReview : AttemptStatus.Completed;

      await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)
        .update({
          EndTime: endTime.toISOString(),
          Score: score,
          Percentage: percentage,
          Passed: passed,
          TimeSpent: timeSpent,
          AnswersJson: JSON.stringify(answers),
          Status: status,
          PointsEarned: score,
          QuestionsAnswered: answers.length,
          QuestionsCorrect: questionsCorrect,
          QuestionsPartial: questionsPartial,
          QuestionsIncorrect: questionsIncorrect,
          QuestionsSkipped: questionsSkipped
        });

      // Update quiz statistics
      await this.updateQuizStatistics(attempt.QuizId);

      // Update question analytics
      await this.updateQuestionAnalytics(answers);

      // Generate certificate if passed and enabled
      let certificateUrl: string | undefined;
      if (passed && quiz.GenerateCertificate) {
        const certificate = await this.generateCertificate(attemptId);
        certificateUrl = certificate?.CertificateUrl;
      }

      // Generate improvement areas
      const improvementAreas = await this.analyzeImprovementAreas(answers, attempt.QuizId);

      logger.info("QuizService", `Quiz attempt submitted: ${attemptId}, Score: ${percentage}%`);

      return {
        attemptId,
        quizId: attempt.QuizId,
        score,
        maxScore,
        percentage,
        passed,
        timeSpent,
        answers,
        pointsEarned: score,
        requiresManualReview,
        pendingQuestions,
        certificateUrl,
        feedback: passed
          ? "Congratulations! You have passed this quiz."
          : "You did not meet the passing score. Please review the material and try again.",
        improvementAreas
      };
    } catch (error) {
      logger.error("QuizService", "Failed to submit quiz attempt", error);
      return null;
    }
  }

  /**
   * Grade a single answer with full support for all question types
   */
  public gradeAnswer(question: IQuizQuestion, userResponse: {
    selectedAnswer?: string;
    selectedAnswers?: string[];
    matchingAnswers?: { left: string; right: string }[];
    orderingAnswers?: string[];
    hotspotCoordinates?: { x: number; y: number };
    essayText?: string;
    ratingValue?: number;
    fillInBlanks?: string[];
  }): IQuizAnswer {
    let isCorrect = false;
    let isPartiallyCorrect = false;
    let pointsEarned = 0;
    let feedback = "";

    switch (question.QuestionType) {
      case QuestionType.MultipleChoice:
      case QuestionType.TrueFalse:
        isCorrect = userResponse.selectedAnswer === question.CorrectAnswer;
        pointsEarned = isCorrect ? question.Points : 0;
        feedback = isCorrect ? question.CorrectFeedback || "" : question.IncorrectFeedback || "";
        break;

      case QuestionType.MultipleSelect:
        const correctAnswers = (question.CorrectAnswers || "").split(";").filter(a => a);
        const userAnswers = userResponse.selectedAnswers || [];

        const correctCount = userAnswers.filter(a => correctAnswers.includes(a)).length;
        const incorrectCount = userAnswers.filter(a => !correctAnswers.includes(a)).length;

        isCorrect = correctCount === correctAnswers.length && incorrectCount === 0;

        if (question.PartialCreditEnabled && !isCorrect && correctCount > 0) {
          isPartiallyCorrect = true;
          const partialScore = (correctCount / correctAnswers.length) * question.Points;
          const penalty = incorrectCount * (question.NegativePoints || 0);
          pointsEarned = Math.max(0, partialScore - penalty);
          feedback = question.PartialFeedback || "";
        } else {
          pointsEarned = isCorrect ? question.Points : 0;
          feedback = isCorrect ? question.CorrectFeedback || "" : question.IncorrectFeedback || "";
        }
        break;

      case QuestionType.ShortAnswer:
        const acceptedAnswers = question.AcceptedAnswers
          ? JSON.parse(question.AcceptedAnswers) as string[]
          : [question.CorrectAnswer];

        const userAnswer = question.CaseSensitive
          ? userResponse.selectedAnswer
          : userResponse.selectedAnswer?.toLowerCase();

        isCorrect = acceptedAnswers.some(accepted => {
          const compareAnswer = question.CaseSensitive ? accepted : accepted.toLowerCase();
          return userAnswer === compareAnswer ||
                 userAnswer?.trim() === compareAnswer.trim();
        });

        pointsEarned = isCorrect ? question.Points : 0;
        feedback = isCorrect ? question.CorrectFeedback || "" : question.IncorrectFeedback || "";
        break;

      case QuestionType.FillInBlank:
        const blanks = question.BlankAnswers ? JSON.parse(question.BlankAnswers) as { position: number; acceptedAnswers: string[] }[] : [];
        const userBlanks = userResponse.fillInBlanks || [];

        let correctBlanks = 0;
        blanks.forEach((blank, index) => {
          const userBlank = question.CaseSensitive
            ? userBlanks[index]
            : userBlanks[index]?.toLowerCase();

          const isBlankCorrect = blank.acceptedAnswers.some(accepted => {
            const compareAnswer = question.CaseSensitive ? accepted : accepted.toLowerCase();
            return userBlank === compareAnswer;
          });

          if (isBlankCorrect) correctBlanks++;
        });

        isCorrect = correctBlanks === blanks.length;
        if (question.PartialCreditEnabled && !isCorrect && correctBlanks > 0) {
          isPartiallyCorrect = true;
          pointsEarned = (correctBlanks / blanks.length) * question.Points;
        } else {
          pointsEarned = isCorrect ? question.Points : 0;
        }
        break;

      case QuestionType.Matching:
        const correctPairs = question.MatchingPairs ? JSON.parse(question.MatchingPairs) as { left: string; right: string }[] : [];
        const userPairs = userResponse.matchingAnswers || [];

        let correctMatches = 0;
        correctPairs.forEach(correct => {
          const userMatch = userPairs.find(u => u.left === correct.left);
          if (userMatch && userMatch.right === correct.right) {
            correctMatches++;
          }
        });

        isCorrect = correctMatches === correctPairs.length;
        if (question.PartialCreditEnabled && !isCorrect && correctMatches > 0) {
          isPartiallyCorrect = true;
          pointsEarned = (correctMatches / correctPairs.length) * question.Points;
        } else {
          pointsEarned = isCorrect ? question.Points : 0;
        }
        break;

      case QuestionType.Ordering:
        const correctOrder = question.OrderingItems
          ? JSON.parse(question.OrderingItems) as { id: string; text: string; correctOrder: number }[]
          : [];
        const userOrder = userResponse.orderingAnswers || [];

        // Check if order matches
        let correctPositions = 0;
        correctOrder.forEach(item => {
          const userPosition = userOrder.indexOf(item.id);
          if (userPosition === item.correctOrder - 1) {
            correctPositions++;
          }
        });

        isCorrect = correctPositions === correctOrder.length;
        if (question.PartialCreditEnabled && !isCorrect && correctPositions > 0) {
          isPartiallyCorrect = true;
          pointsEarned = (correctPositions / correctOrder.length) * question.Points;
        } else {
          pointsEarned = isCorrect ? question.Points : 0;
        }
        break;

      case QuestionType.RatingScale:
        const correctRating = question.CorrectRating || 0;
        const tolerance = question.RatingTolerance || 0;
        const userRating = userResponse.ratingValue || 0;

        isCorrect = Math.abs(userRating - correctRating) <= tolerance;

        if (question.PartialCreditEnabled && !isCorrect) {
          const distance = Math.abs(userRating - correctRating);
          const maxDistance = (question.ScaleMax || 5) - (question.ScaleMin || 1);
          if (distance < maxDistance) {
            isPartiallyCorrect = true;
            pointsEarned = ((maxDistance - distance) / maxDistance) * question.Points;
          }
        } else {
          pointsEarned = isCorrect ? question.Points : 0;
        }
        break;

      case QuestionType.ImageChoice:
        isCorrect = userResponse.selectedAnswer === question.CorrectAnswer;
        pointsEarned = isCorrect ? question.Points : 0;
        break;

      case QuestionType.Hotspot:
        const hotspotData = question.HotspotData
          ? JSON.parse(question.HotspotData) as { imageUrl: string; regions: { x: number; y: number; width: number; height: number; isCorrect: boolean }[] }
          : null;

        if (hotspotData && userResponse.hotspotCoordinates) {
          const { x, y } = userResponse.hotspotCoordinates;
          const clickedRegion = hotspotData.regions.find(region =>
            x >= region.x && x <= region.x + region.width &&
            y >= region.y && y <= region.y + region.height
          );

          isCorrect = clickedRegion?.isCorrect ?? false;
          pointsEarned = isCorrect ? question.Points : 0;
        }
        break;

      case QuestionType.Essay:
        // Essays require manual grading
        pointsEarned = 0;
        feedback = "This question requires manual review.";
        break;
    }

    // Apply negative marking
    if (!isCorrect && !isPartiallyCorrect && question.NegativeMarking && question.NegativePoints) {
      pointsEarned = -question.NegativePoints;
    }

    return {
      questionId: question.Id,
      questionType: question.QuestionType,
      selectedAnswer: userResponse.selectedAnswer || "",
      selectedAnswers: userResponse.selectedAnswers,
      matchingAnswers: userResponse.matchingAnswers,
      orderingAnswers: userResponse.orderingAnswers,
      hotspotCoordinates: userResponse.hotspotCoordinates,
      essayText: userResponse.essayText,
      ratingValue: userResponse.ratingValue,
      isCorrect,
      isPartiallyCorrect,
      pointsEarned: Math.round(pointsEarned * 100) / 100,
      maxPoints: question.Points,
      feedback: feedback || question.Explanation
    };
  }

  /**
   * Manual grade for essay questions
   */
  public async manualGradeQuestion(
    attemptId: number,
    questionId: number,
    grade: number,
    feedback: string,
    reviewerId: number
  ): Promise<void> {
    try {
      const attempt = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)() as IQuizAttempt;

      const answers: IQuizAnswer[] = attempt.AnswersJson
        ? JSON.parse(attempt.AnswersJson)
        : [];

      // Update the specific answer
      const answerIndex = answers.findIndex(a => a.questionId === questionId);
      if (answerIndex >= 0) {
        answers[answerIndex].manualGrade = grade;
        answers[answerIndex].manualFeedback = feedback;
        answers[answerIndex].pointsEarned = grade;
        answers[answerIndex].isCorrect = grade > 0;
      }

      // Recalculate total score
      const newScore = answers.reduce((sum, a) => sum + a.pointsEarned, 0);
      const percentage = attempt.MaxScore > 0
        ? Math.round((newScore / attempt.MaxScore) * 100)
        : 0;

      // Check if all essays are graded
      const pendingEssays = answers.filter(
        a => a.questionType === QuestionType.Essay && a.manualGrade === undefined
      );

      const quiz = await this.getQuizById(attempt.QuizId);
      const passed = percentage >= (quiz?.PassingScore || 70);

      await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)
        .update({
          Score: newScore,
          Percentage: percentage,
          Passed: passed,
          AnswersJson: JSON.stringify(answers),
          Status: pendingEssays.length === 0 ? AttemptStatus.Completed : AttemptStatus.PendingReview,
          ReviewedById: reviewerId,
          ReviewedDate: new Date().toISOString()
        });

      // Generate certificate if now passed
      if (passed && quiz?.GenerateCertificate && pendingEssays.length === 0) {
        await this.generateCertificate(attemptId);
      }

      logger.info("QuizService", `Manual grade applied: attempt ${attemptId}, question ${questionId}`);
    } catch (error) {
      logger.error("QuizService", "Failed to apply manual grade", error);
      throw error;
    }
  }

  /**
   * Get user's quiz attempts
   */
  public async getUserQuizAttempts(quizId: number, userId: number): Promise<IQuizAttempt[]> {
    try {
      const attempts = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.filter(`QuizId eq ${quizId} and UserId eq ${userId}`)
        .orderBy("AttemptNumber", false)();

      return attempts as IQuizAttempt[];
    } catch (error) {
      logger.error("QuizService", "Failed to get user quiz attempts", error);
      return [];
    }
  }

  /**
   * Get attempt by ID
   */
  public async getAttemptById(attemptId: number): Promise<IQuizAttempt | null> {
    try {
      const attempt = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)();

      return attempt as IQuizAttempt;
    } catch (error) {
      logger.error("QuizService", `Failed to get attempt ${attemptId}`, error);
      return null;
    }
  }

  /**
   * Abandon quiz attempt
   */
  public async abandonQuizAttempt(attemptId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)
        .update({
          Status: AttemptStatus.Abandoned,
          EndTime: new Date().toISOString()
        });

      logger.info("QuizService", `Quiz attempt abandoned: ${attemptId}`);
    } catch (error) {
      logger.error("QuizService", "Failed to abandon quiz attempt", error);
      throw error;
    }
  }

  /**
   * Check if user can take quiz
   */
  public async canUserTakeQuiz(quizId: number, userId: number): Promise<{
    canTake: boolean;
    reason?: string;
    attemptsRemaining?: number;
    nextAvailableDate?: Date;
  }> {
    try {
      const quiz = await this.getQuizById(quizId);

      if (!quiz || !quiz.IsActive) {
        return { canTake: false, reason: "Quiz is not available" };
      }

      // Check status
      if (quiz.Status === QuizStatus.Draft) {
        return { canTake: false, reason: "Quiz is not published" };
      }

      if (quiz.Status === QuizStatus.Archived) {
        return { canTake: false, reason: "Quiz has been archived" };
      }

      // Check scheduling
      const now = new Date();
      if (quiz.ScheduledStartDate && new Date(quiz.ScheduledStartDate) > now) {
        return {
          canTake: false,
          reason: "Quiz has not started yet",
          nextAvailableDate: new Date(quiz.ScheduledStartDate)
        };
      }

      if (quiz.ScheduledEndDate && new Date(quiz.ScheduledEndDate) < now) {
        return { canTake: false, reason: "Quiz has ended" };
      }

      // Check attempts
      const attempts = await this.getUserQuizAttempts(quizId, userId);
      const completedAttempts = attempts.filter(
        a => a.Status === AttemptStatus.Completed || a.Status === AttemptStatus.PendingReview
      );
      const attemptsRemaining = quiz.MaxAttempts - completedAttempts.length;

      if (attemptsRemaining <= 0) {
        return { canTake: false, reason: "Maximum attempts reached", attemptsRemaining: 0 };
      }

      // Check for in-progress attempt
      const inProgressAttempt = attempts.find(a => a.Status === AttemptStatus.InProgress);
      if (inProgressAttempt) {
        return {
          canTake: false,
          reason: "You have an in-progress attempt",
          attemptsRemaining
        };
      }

      return { canTake: true, attemptsRemaining };
    } catch (error) {
      logger.error("QuizService", "Failed to check quiz eligibility", error);
      return { canTake: false, reason: "Error checking eligibility" };
    }
  }

  // ============================================================================
  // Question Banks
  // ============================================================================

  /**
   * Get all question banks
   */
  public async getQuestionBanks(): Promise<IQuestionBank[]> {
    try {
      const banks = await this.sp.web.lists
        .getByTitle(this.bankListName)
        .items.orderBy("Title", true)();

      return banks as IQuestionBank[];
    } catch (error) {
      logger.error("QuizService", "Failed to get question banks", error);
      return [];
    }
  }

  /**
   * Create question bank
   */
  public async createQuestionBank(bank: Partial<IQuestionBank>): Promise<IQuestionBank | null> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.bankListName)
        .items.add({
          Title: bank.Title,
          Description: bank.Description,
          Category: bank.Category,
          Tags: bank.Tags,
          QuestionCount: 0,
          IsPublic: bank.IsPublic ?? false
        });

      return result.data as IQuestionBank;
    } catch (error) {
      logger.error("QuizService", "Failed to create question bank", error);
      return null;
    }
  }

  /**
   * Add questions to bank
   */
  public async addQuestionsToBank(
    bankId: number,
    questionIds: number[]
  ): Promise<void> {
    try {
      for (const questionId of questionIds) {
        await this.updateQuestion(questionId, { QuestionBankId: bankId });
      }

      // Update bank question count
      const questions = await this.getQuestionsFromBank(bankId);
      await this.sp.web.lists
        .getByTitle(this.bankListName)
        .items.getById(bankId)
        .update({ QuestionCount: questions.length });

    } catch (error) {
      logger.error("QuizService", "Failed to add questions to bank", error);
      throw error;
    }
  }

  // ============================================================================
  // Quiz Sections
  // ============================================================================

  /**
   * Get sections for a quiz
   */
  public async getQuizSections(quizId: number): Promise<IQuizSection[]> {
    try {
      const sections = await this.sp.web.lists
        .getByTitle(this.sectionListName)
        .items.filter(`QuizId eq ${quizId}`)
        .orderBy("Order", true)();

      return sections as IQuizSection[];
    } catch (error) {
      logger.error("QuizService", "Failed to get quiz sections", error);
      return [];
    }
  }

  /**
   * Create quiz section
   */
  public async createSection(section: Partial<IQuizSection>): Promise<IQuizSection | null> {
    try {
      const existingSections = await this.getQuizSections(section.QuizId!);

      const result = await this.sp.web.lists
        .getByTitle(this.sectionListName)
        .items.add({
          Title: section.Title,
          QuizId: section.QuizId,
          Description: section.Description,
          Order: section.Order || existingSections.length + 1,
          RandomizeWithinSection: section.RandomizeWithinSection ?? false,
          QuestionsRequired: section.QuestionsRequired
        });

      return result.data as IQuizSection;
    } catch (error) {
      logger.error("QuizService", "Failed to create section", error);
      return null;
    }
  }

  // ============================================================================
  // Statistics & Analytics
  // ============================================================================

  /**
   * Get comprehensive quiz statistics
   */
  public async getQuizStatistics(quizId: number): Promise<IQuizStatistics> {
    try {
      const attempts = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.filter(`QuizId eq ${quizId}`)() as IQuizAttempt[];

      const completedAttempts = attempts.filter(
        a => a.Status === AttemptStatus.Completed || a.Status === AttemptStatus.PendingReview
      );
      const passedAttempts = completedAttempts.filter(a => a.Passed);
      const failedAttempts = completedAttempts.filter(a => !a.Passed);
      const pendingReview = attempts.filter(a => a.Status === AttemptStatus.PendingReview);

      // Calculate unique users
      const uniqueUsers = new Set(attempts.map(a => a.UserId)).size;

      // Calculate score statistics
      const scores = completedAttempts.map(a => a.Percentage).sort((a, b) => a - b);
      const averageScore = scores.length > 0
        ? scores.reduce((sum, s) => sum + s, 0) / scores.length
        : 0;
      const medianScore = scores.length > 0
        ? scores[Math.floor(scores.length / 2)]
        : 0;
      const highestScore = scores.length > 0 ? Math.max(...scores) : 0;
      const lowestScore = scores.length > 0 ? Math.min(...scores) : 0;

      // Calculate time statistics
      const times = completedAttempts.map(a => a.TimeSpent || 0);
      const averageTimeSpent = times.length > 0
        ? times.reduce((sum, t) => sum + t, 0) / times.length
        : 0;

      // Calculate rates
      const completionRate = attempts.length > 0
        ? (completedAttempts.length / attempts.length) * 100
        : 0;
      const passRate = completedAttempts.length > 0
        ? (passedAttempts.length / completedAttempts.length) * 100
        : 0;

      // Score distribution
      const scoreDistribution = [
        { range: "0-20%", count: scores.filter(s => s >= 0 && s < 20).length },
        { range: "20-40%", count: scores.filter(s => s >= 20 && s < 40).length },
        { range: "40-60%", count: scores.filter(s => s >= 40 && s < 60).length },
        { range: "60-80%", count: scores.filter(s => s >= 60 && s < 80).length },
        { range: "80-100%", count: scores.filter(s => s >= 80 && s <= 100).length }
      ];

      // Get question analytics
      const questionAnalytics = await this.getQuestionAnalytics(quizId);

      return {
        totalAttempts: attempts.length,
        uniqueUsers,
        passedAttempts: passedAttempts.length,
        failedAttempts: failedAttempts.length,
        pendingReview: pendingReview.length,
        averageScore: Math.round(averageScore),
        medianScore,
        highestScore,
        lowestScore,
        averageTimeSpent: Math.round(averageTimeSpent),
        completionRate: Math.round(completionRate),
        passRate: Math.round(passRate),
        averageAttemptsPerUser: uniqueUsers > 0 ? attempts.length / uniqueUsers : 0,
        scoreDistribution,
        questionAnalytics
      };
    } catch (error) {
      logger.error("QuizService", "Failed to get quiz statistics", error);
      return {
        totalAttempts: 0,
        uniqueUsers: 0,
        passedAttempts: 0,
        failedAttempts: 0,
        pendingReview: 0,
        averageScore: 0,
        medianScore: 0,
        highestScore: 0,
        lowestScore: 0,
        averageTimeSpent: 0,
        completionRate: 0,
        passRate: 0,
        averageAttemptsPerUser: 0,
        scoreDistribution: [],
        questionAnalytics: []
      };
    }
  }

  /**
   * Get per-question analytics
   */
  private async getQuestionAnalytics(quizId: number): Promise<IQuestionAnalytics[]> {
    try {
      const questions = await this.getQuizQuestions(quizId);
      const attempts = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.filter(`QuizId eq ${quizId} and Status eq 'Completed'`)
        .select("AnswersJson")() as { AnswersJson: string }[];

      const analytics: IQuestionAnalytics[] = [];

      for (const question of questions) {
        let timesAnswered = 0;
        let correctCount = 0;
        let partialCount = 0;
        let incorrectCount = 0;
        let skippedCount = 0;
        let totalTime = 0;
        let totalScore = 0;
        const wrongAnswers: Map<string, number> = new Map();

        for (const attempt of attempts) {
          const answers: IQuizAnswer[] = attempt.AnswersJson
            ? JSON.parse(attempt.AnswersJson)
            : [];
          const answer = answers.find(a => a.questionId === question.Id);

          if (answer) {
            timesAnswered++;
            totalScore += answer.pointsEarned;
            if (answer.timeSpent) totalTime += answer.timeSpent;

            if (answer.isCorrect) {
              correctCount++;
            } else if (answer.isPartiallyCorrect) {
              partialCount++;
            } else if (answer.selectedAnswer || answer.selectedAnswers?.length) {
              incorrectCount++;
              const wrongAnswer = answer.selectedAnswer || answer.selectedAnswers?.join(", ") || "";
              wrongAnswers.set(wrongAnswer, (wrongAnswers.get(wrongAnswer) || 0) + 1);
            } else {
              skippedCount++;
            }
          }
        }

        // Calculate discrimination index (simplified)
        const discriminationIndex = timesAnswered > 0
          ? (correctCount - incorrectCount) / timesAnswered
          : 0;

        // Difficulty index (% correct)
        const difficultyIndex = timesAnswered > 0
          ? correctCount / timesAnswered
          : 0;

        analytics.push({
          questionId: question.Id,
          questionText: question.QuestionText.substring(0, 100),
          questionType: question.QuestionType,
          timesAnswered,
          correctRate: timesAnswered > 0 ? (correctCount / timesAnswered) * 100 : 0,
          partialRate: timesAnswered > 0 ? (partialCount / timesAnswered) * 100 : 0,
          incorrectRate: timesAnswered > 0 ? (incorrectCount / timesAnswered) * 100 : 0,
          skippedRate: timesAnswered > 0 ? (skippedCount / timesAnswered) * 100 : 0,
          averageTime: timesAnswered > 0 ? totalTime / timesAnswered : 0,
          averageScore: timesAnswered > 0 ? totalScore / timesAnswered : 0,
          discriminationIndex,
          difficultyIndex,
          commonWrongAnswers: Array.from(wrongAnswers.entries())
            .map(([answer, count]) => ({ answer, count }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 5)
        });
      }

      return analytics;
    } catch (error) {
      logger.error("QuizService", "Failed to get question analytics", error);
      return [];
    }
  }

  /**
   * Update quiz statistics (called after attempt submission)
   */
  private async updateQuizStatistics(quizId: number): Promise<void> {
    try {
      const stats = await this.getQuizStatistics(quizId);

      await this.sp.web.lists
        .getByTitle(this.quizListName)
        .items.getById(quizId)
        .update({
          AverageScore: stats.averageScore,
          CompletionRate: stats.completionRate
        });
    } catch (error) {
      logger.error("QuizService", "Failed to update quiz statistics", error);
    }
  }

  /**
   * Update question-level analytics
   */
  private async updateQuestionAnalytics(answers: IQuizAnswer[]): Promise<void> {
    try {
      for (const answer of answers) {
        const question = await this.getQuestionById(answer.questionId);
        if (!question) continue;

        const timesAnswered = (question.TimesAnswered || 0) + 1;
        const timesCorrect = (question.TimesCorrect || 0) + (answer.isCorrect ? 1 : 0);
        const avgTime = question.AverageTime || 0;
        const newAvgTime = answer.timeSpent
          ? ((avgTime * (timesAnswered - 1)) + answer.timeSpent) / timesAnswered
          : avgTime;

        await this.updateQuestion(answer.questionId, {
          TimesAnswered: timesAnswered,
          TimesCorrect: timesCorrect,
          AverageTime: newAvgTime
        });
      }
    } catch (error) {
      logger.error("QuizService", "Failed to update question analytics", error);
    }
  }

  /**
   * Analyze improvement areas based on answers
   */
  private async analyzeImprovementAreas(
    answers: IQuizAnswer[],
    quizId: number
  ): Promise<string[]> {
    const incorrectAnswers = answers.filter(a => !a.isCorrect && !a.isPartiallyCorrect);
    const questions = await this.getQuizQuestions(quizId);

    const categoryMisses: Map<string, number> = new Map();

    for (const answer of incorrectAnswers) {
      const question = questions.find(q => q.Id === answer.questionId);
      if (question?.Category) {
        categoryMisses.set(
          question.Category,
          (categoryMisses.get(question.Category) || 0) + 1
        );
      }
    }

    return Array.from(categoryMisses.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([category]) => category);
  }

  // ============================================================================
  // Certificate Generation
  // ============================================================================

  /**
   * Generate certificate for passed attempt
   */
  public async generateCertificate(attemptId: number): Promise<ICertificate | null> {
    try {
      const attempt = await this.getAttemptById(attemptId);
      if (!attempt || !attempt.Passed) {
        throw new Error("Cannot generate certificate for non-passing attempt");
      }

      const quiz = await this.getQuizById(attempt.QuizId);
      if (!quiz?.GenerateCertificate) {
        throw new Error("Quiz does not have certificate generation enabled");
      }

      // Generate unique certificate number
      const certificateNumber = `CERT-${quiz.Id}-${attemptId}-${Date.now()}`;

      // Create certificate record
      const result = await this.sp.web.lists
        .getByTitle(this.certificateListName)
        .items.add({
          Title: certificateNumber,
          AttemptId: attemptId,
          UserId: attempt.UserId,
          UserName: attempt.UserName,
          QuizTitle: quiz.Title,
          Score: attempt.Percentage,
          PassedDate: attempt.EndTime,
          CertificateNumber: certificateNumber
        });

      // Update attempt with certificate info
      await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.getById(attemptId)
        .update({
          CertificateGenerated: true,
          CertificateUrl: `/certificates/${certificateNumber}`
        });

      logger.info("QuizService", `Certificate generated: ${certificateNumber}`);

      return result.data as ICertificate;
    } catch (error) {
      logger.error("QuizService", "Failed to generate certificate", error);
      return null;
    }
  }

  /**
   * Get certificate by ID
   */
  public async getCertificate(certificateId: number): Promise<ICertificate | null> {
    try {
      const cert = await this.sp.web.lists
        .getByTitle(this.certificateListName)
        .items.getById(certificateId)();

      return cert as ICertificate;
    } catch (error) {
      logger.error("QuizService", "Failed to get certificate", error);
      return null;
    }
  }

  /**
   * Get user's certificates
   */
  public async getUserCertificates(userId: number): Promise<ICertificate[]> {
    try {
      const certs = await this.sp.web.lists
        .getByTitle(this.certificateListName)
        .items.filter(`UserId eq ${userId}`)
        .orderBy("PassedDate", false)();

      return certs as ICertificate[];
    } catch (error) {
      logger.error("QuizService", "Failed to get user certificates", error);
      return [];
    }
  }

  // ============================================================================
  // Import/Export
  // ============================================================================

  /**
   * Export quiz with questions
   */
  public async exportQuiz(quizId: number): Promise<IQuizExportData> {
    const quiz = await this.getQuizById(quizId);
    const questions = await this.getQuizQuestions(quizId);
    const sections = await this.getQuizSections(quizId);

    if (!quiz) throw new Error("Quiz not found");

    // Clean up data for export
    const cleanQuiz = { ...quiz };
    delete (cleanQuiz as any).Id;
    delete cleanQuiz.AverageScore;
    delete cleanQuiz.CompletionRate;

    const cleanQuestions = questions.map(q => {
      const clean = { ...q };
      delete (clean as any).Id;
      delete clean.QuizId;
      delete clean.TimesAnswered;
      delete clean.TimesCorrect;
      delete clean.AverageTime;
      return clean;
    });

    return {
      version: "1.0",
      exportDate: new Date().toISOString(),
      quiz: cleanQuiz,
      sections,
      questions: cleanQuestions
    };
  }

  /**
   * Import quiz from export data
   */
  public async importQuiz(
    data: IQuizExportData,
    options?: {
      newTitle?: string;
      policyId?: number;
      asDraft?: boolean;
    }
  ): Promise<IQuiz | null> {
    try {
      // Create quiz
      const quizData = {
        ...data.quiz,
        Title: options?.newTitle || data.quiz.Title,
        PolicyId: options?.policyId || data.quiz.PolicyId,
        Status: options?.asDraft ? QuizStatus.Draft : data.quiz.Status
      };

      const quiz = await this.createQuiz(quizData);
      if (!quiz) throw new Error("Failed to create quiz");

      // Create sections
      const sectionMap = new Map<number, number>();
      for (const section of data.sections) {
        const newSection = await this.createSection({
          ...section,
          QuizId: quiz.Id
        });
        if (newSection && section.Id) {
          sectionMap.set(section.Id, newSection.Id);
        }
      }

      // Create questions
      for (const question of data.questions) {
        await this.createQuestion({
          ...question,
          QuizId: quiz.Id,
          SectionId: question.SectionId ? sectionMap.get(question.SectionId) : undefined
        });
      }

      logger.info("QuizService", `Quiz imported: ${quiz.Title}`);
      return quiz;
    } catch (error) {
      logger.error("QuizService", "Failed to import quiz", error);
      return null;
    }
  }

  /**
   * Import questions from CSV
   */
  public async importQuestionsFromCSV(
    quizId: number,
    csvData: string
  ): Promise<{ imported: number; errors: string[] }> {
    const errors: string[] = [];
    let imported = 0;

    try {
      const lines = csvData.split("\n").slice(1); // Skip header

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;

        try {
          const [
            questionText,
            questionType,
            optionA,
            optionB,
            optionC,
            optionD,
            correctAnswer,
            explanation,
            points,
            difficulty
          ] = this.parseCSVLine(line);

          await this.createQuestion({
            QuizId: quizId,
            QuestionText: questionText,
            QuestionType: questionType as QuestionType || QuestionType.MultipleChoice,
            OptionA: optionA,
            OptionB: optionB,
            OptionC: optionC,
            OptionD: optionD,
            CorrectAnswer: correctAnswer,
            Explanation: explanation,
            Points: parseInt(points) || 10,
            DifficultyLevel: difficulty as DifficultyLevel || DifficultyLevel.Medium
          });

          imported++;
        } catch (lineError) {
          errors.push(`Line ${i + 2}: ${lineError}`);
        }
      }

      return { imported, errors };
    } catch (error) {
      logger.error("QuizService", "Failed to import questions from CSV", error);
      return { imported, errors: [`Import failed: ${error}`] };
    }
  }

  /**
   * Export questions to CSV
   */
  public exportQuestionsToCSV(questions: IQuizQuestion[]): string {
    const header = [
      "QuestionText",
      "QuestionType",
      "OptionA",
      "OptionB",
      "OptionC",
      "OptionD",
      "CorrectAnswer",
      "Explanation",
      "Points",
      "Difficulty"
    ].join(",");

    const rows = questions.map(q => [
      this.escapeCSV(q.QuestionText),
      q.QuestionType,
      this.escapeCSV(q.OptionA || ""),
      this.escapeCSV(q.OptionB || ""),
      this.escapeCSV(q.OptionC || ""),
      this.escapeCSV(q.OptionD || ""),
      this.escapeCSV(q.CorrectAnswer),
      this.escapeCSV(q.Explanation || ""),
      q.Points.toString(),
      q.DifficultyLevel
    ].join(","));

    return [header, ...rows].join("\n");
  }

  // ============================================================================
  // User Progress & Summary
  // ============================================================================

  /**
   * Get user's quiz history
   */
  public async getUserQuizHistory(userId: number): Promise<IQuizAttempt[]> {
    try {
      const attempts = await this.sp.web.lists
        .getByTitle(this.attemptListName)
        .items.filter(`UserId eq ${userId} and (Status eq 'Completed' or Status eq 'Pending Review')`)
        .orderBy("EndTime", false)
        .top(100)();

      return attempts as IQuizAttempt[];
    } catch (error) {
      logger.error("QuizService", "Failed to get user quiz history", error);
      return [];
    }
  }

  /**
   * Get quiz summary for policy view
   */
  public async getQuizSummary(policyId: number, userId: number): Promise<{
    hasQuiz: boolean;
    quiz?: IQuiz;
    attempts: number;
    bestScore: number;
    passed: boolean;
    canRetake: boolean;
    certificateUrl?: string;
  }> {
    try {
      const quizzes = await this.getQuizzesByPolicy(policyId);

      if (quizzes.length === 0) {
        return { hasQuiz: false, attempts: 0, bestScore: 0, passed: false, canRetake: false };
      }

      const quiz = quizzes[0];
      const attempts = await this.getUserQuizAttempts(quiz.Id, userId);
      const completedAttempts = attempts.filter(
        a => a.Status === AttemptStatus.Completed || a.Status === AttemptStatus.PendingReview
      );

      const bestScore = completedAttempts.length > 0
        ? Math.max(...completedAttempts.map(a => a.Percentage))
        : 0;

      const passedAttempt = completedAttempts.find(a => a.Passed);
      const passed = !!passedAttempt;
      const canRetake = completedAttempts.length < quiz.MaxAttempts;

      return {
        hasQuiz: true,
        quiz,
        attempts: completedAttempts.length,
        bestScore,
        passed,
        canRetake,
        certificateUrl: passedAttempt?.CertificateUrl
      };
    } catch (error) {
      logger.error("QuizService", "Failed to get quiz summary", error);
      return { hasQuiz: false, attempts: 0, bestScore: 0, passed: false, canRetake: false };
    }
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Shuffle array (Fisher-Yates algorithm)
   */
  private shuffleArray<T>(array: T[]): T[] {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
  }

  /**
   * Parse CSV line handling quoted values
   */
  private parseCSVLine(line: string): string[] {
    const result: string[] = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];

      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === "," && !inQuotes) {
        result.push(current);
        current = "";
      } else {
        current += char;
      }
    }

    result.push(current);
    return result;
  }

  /**
   * Escape value for CSV
   */
  private escapeCSV(value: string): string {
    if (value.includes(",") || value.includes('"') || value.includes("\n")) {
      return `"${value.replace(/"/g, '""')}"`;
    }
    return value;
  }
}
