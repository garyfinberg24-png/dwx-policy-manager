// @ts-nocheck
// SkillsCompetenciesService - Skills & Competencies Data Access Layer
// Handles all SharePoint list operations for Skills & Competencies module

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';

import {
  ISkill,
  IUserSkill,
  IRoleCompetency,
  IRoleSkillRequirement,
  ICareerStep,
  ISkillGap,
  ISkillStrength,
  ISkillsGapAnalysis,
  ProficiencyLevel,
  SkillDomain,
  SkillSource
} from '../models/ITraining';
import { logger } from './LoggingService';

// ============================================================================
// LIST INTERFACES - Matching SharePoint List Schema
// ============================================================================

/** SharePoint list item for JML_SkillCategories */
export interface ISkillCategoryListItem {
  Id: number;
  Title: string;
  Description?: string;
  Icon?: string;
  SortOrder?: number;
  ParentCategoryId?: number;
}

/** SharePoint list item for JML_Skills */
export interface ISkillListItem {
  Id: number;
  Title: string;
  SkillCode: string;
  Description?: string;
  Domain: string;
  SkillCategoryId?: number;
  IsCore: boolean;
  SkillStatus: string;
  Tags?: string;
}

/** SharePoint list item for JML_UserSkills */
export interface IUserSkillListItem {
  Id: number;
  Title: string;
  UserId: number;
  UserEmail: string;
  SkillId: number;
  SkillCode?: string;
  SelfRating?: number;
  ManagerRating?: number;
  VerifiedRating?: number;
  LastAssessedDate?: string;
  AssessedBy?: string;
  Evidence?: string;
  SkillSource: string;
  Notes?: string;
}

/** SharePoint list item for JML_RoleCompetencies */
export interface IRoleCompetencyListItem {
  Id: number;
  Title: string;
  RoleCode: string;
  Level: string;
  Department?: string;
  RoleStatus: string;
  RequiredSkills?: string;
  PreferredSkills?: string;
  SuccessionPath?: string;
}

/** SharePoint list item for JML_SkillsAssessments */
export interface ISkillsAssessmentListItem {
  Id: number;
  Title: string;
  UserId: number;
  UserEmail: string;
  RoleId?: number;
  AssessmentDate: string;
  OverallScore?: number;
  AssessmentStatus: string;
  AssessorId?: number;
  AssessorName?: string;
  Comments?: string;
}

/** SharePoint list item for JML_SkillTests */
export interface ISkillTestListItem {
  Id: number;
  Title: string;
  SkillId: number;
  SkillCode?: string;
  Description?: string;
  Questions?: string;
  PassingScore: number;
  TimeLimit?: number;
  DifficultyLevel: string;
  TestStatus: string;
}

/** SharePoint list item for JML_SkillTestAttempts */
export interface ISkillTestAttemptListItem {
  Id: number;
  Title: string;
  TestId: number;
  UserId: number;
  UserEmail: string;
  AttemptNumber: number;
  Score?: number;
  Passed?: boolean;
  StartTime: string;
  EndTime?: string;
  Answers?: string;
}

// ============================================================================
// SERVICE CLASS
// ============================================================================

export class SkillsCompetenciesService {
  private sp: SPFI;

  // List names
  private readonly SKILL_CATEGORIES_LIST = 'JML_SkillCategories';
  private readonly SKILLS_LIST = 'JML_Skills';
  private readonly USER_SKILLS_LIST = 'JML_UserSkills';
  private readonly ROLE_COMPETENCIES_LIST = 'JML_RoleCompetencies';
  private readonly SKILLS_ASSESSMENTS_LIST = 'JML_SkillsAssessments';
  private readonly SKILL_TESTS_LIST = 'JML_SkillTests';
  private readonly SKILL_TEST_ATTEMPTS_LIST = 'JML_SkillTestAttempts';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ===================================================================
  // SKILL CATEGORIES
  // ===================================================================

  /**
   * Get all skill categories
   */
  public async getSkillCategories(): Promise<ISkillCategoryListItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SKILL_CATEGORIES_LIST).items
        .select('Id', 'Title', 'Description', 'Icon', 'SortOrder', 'ParentCategoryId')
        .orderBy('SortOrder', true)
        .top(500)();

      return items as ISkillCategoryListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error fetching skill categories', error);
      throw error;
    }
  }

  /**
   * Get skill category by ID
   */
  public async getSkillCategoryById(id: number): Promise<ISkillCategoryListItem> {
    try {
      const item = await this.sp.web.lists.getByTitle(this.SKILL_CATEGORIES_LIST).items
        .getById(id)
        .select('Id', 'Title', 'Description', 'Icon', 'SortOrder', 'ParentCategoryId')();

      return item as ISkillCategoryListItem;
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching skill category ${id}`, error);
      throw error;
    }
  }

  // ===================================================================
  // SKILLS
  // ===================================================================

  /**
   * Get all skills
   */
  public async getSkills(filter?: string): Promise<ISkill[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.SKILLS_LIST).items
        .select('Id', 'Title', 'SkillCode', 'Description', 'Domain', 'SkillCategoryId', 'IsCore', 'SkillStatus', 'Tags')
        .orderBy('Title', true)
        .top(1000);

      if (filter) {
        query = query.filter(filter);
      }

      const items = await query() as ISkillListItem[];
      return items.map(this.mapSkillListItemToSkill);
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error fetching skills', error);
      throw error;
    }
  }

  /**
   * Get active skills only
   */
  public async getActiveSkills(): Promise<ISkill[]> {
    return this.getSkills("SkillStatus eq 'Active'");
  }

  /**
   * Get skill by ID
   */
  public async getSkillById(id: number): Promise<ISkill> {
    try {
      const item = await this.sp.web.lists.getByTitle(this.SKILLS_LIST).items
        .getById(id)
        .select('Id', 'Title', 'SkillCode', 'Description', 'Domain', 'SkillCategoryId', 'IsCore', 'SkillStatus', 'Tags')() as ISkillListItem;

      return this.mapSkillListItemToSkill(item);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching skill ${id}`, error);
      throw error;
    }
  }

  /**
   * Get skills by domain
   */
  public async getSkillsByDomain(domain: SkillDomain): Promise<ISkill[]> {
    return this.getSkills(`Domain eq '${domain}'`);
  }

  /**
   * Get skills by category ID
   */
  public async getSkillsByCategory(categoryId: number): Promise<ISkill[]> {
    return this.getSkills(`SkillCategoryId eq ${categoryId}`);
  }

  /**
   * Search skills by name or code
   */
  public async searchSkills(query: string): Promise<ISkill[]> {
    const filter = `substringof('${query}', Title) or substringof('${query}', SkillCode)`;
    return this.getSkills(filter);
  }

  /**
   * Create a new skill
   */
  public async createSkill(skill: {
    title: string;
    skillCode: string;
    description?: string;
    domain: string;
    categoryId?: number;
    isCore: boolean;
    tags?: string;
  }): Promise<ISkill> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.SKILLS_LIST).items.add({
        Title: skill.title,
        SkillCode: skill.skillCode,
        Description: skill.description,
        Domain: skill.domain,
        SkillCategoryId: skill.categoryId,
        IsCore: skill.isCore,
        SkillStatus: 'Active',
        Tags: skill.tags
      });

      return await this.getSkillById(result.data.Id);
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error creating skill', error);
      throw error;
    }
  }

  /**
   * Update an existing skill
   */
  public async updateSkill(id: number, skill: {
    title?: string;
    skillCode?: string;
    description?: string;
    domain?: string;
    categoryId?: number;
    isCore?: boolean;
    status?: string;
    tags?: string;
  }): Promise<ISkill> {
    try {
      const updates: Record<string, unknown> = {};

      if (skill.title !== undefined) updates.Title = skill.title;
      if (skill.skillCode !== undefined) updates.SkillCode = skill.skillCode;
      if (skill.description !== undefined) updates.Description = skill.description;
      if (skill.domain !== undefined) updates.Domain = skill.domain;
      if (skill.categoryId !== undefined) updates.SkillCategoryId = skill.categoryId;
      if (skill.isCore !== undefined) updates.IsCore = skill.isCore;
      if (skill.status !== undefined) updates.SkillStatus = skill.status;
      if (skill.tags !== undefined) updates.Tags = skill.tags;

      await this.sp.web.lists.getByTitle(this.SKILLS_LIST).items
        .getById(id)
        .update(updates);

      return await this.getSkillById(id);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error updating skill ${id}`, error);
      throw error;
    }
  }

  /**
   * Delete a skill (sets status to Deprecated)
   */
  public async deleteSkill(id: number, hardDelete: boolean = false): Promise<void> {
    try {
      if (hardDelete) {
        await this.sp.web.lists.getByTitle(this.SKILLS_LIST).items
          .getById(id)
          .delete();
      } else {
        // Soft delete - set status to Deprecated
        await this.sp.web.lists.getByTitle(this.SKILLS_LIST).items
          .getById(id)
          .update({ SkillStatus: 'Deprecated' });
      }
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error deleting skill ${id}`, error);
      throw error;
    }
  }

  // ===================================================================
  // SKILL CATEGORIES - CRUD
  // ===================================================================

  /**
   * Create a new skill category
   */
  public async createSkillCategory(category: {
    title: string;
    description?: string;
    icon?: string;
    sortOrder?: number;
    parentCategoryId?: number;
  }): Promise<ISkillCategoryListItem> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.SKILL_CATEGORIES_LIST).items.add({
        Title: category.title,
        Description: category.description,
        Icon: category.icon,
        SortOrder: category.sortOrder || 0,
        ParentCategoryId: category.parentCategoryId
      });

      return await this.getSkillCategoryById(result.data.Id);
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error creating skill category', error);
      throw error;
    }
  }

  /**
   * Update an existing skill category
   */
  public async updateSkillCategory(id: number, category: {
    title?: string;
    description?: string;
    icon?: string;
    sortOrder?: number;
    parentCategoryId?: number;
  }): Promise<ISkillCategoryListItem> {
    try {
      const updates: Record<string, unknown> = {};

      if (category.title !== undefined) updates.Title = category.title;
      if (category.description !== undefined) updates.Description = category.description;
      if (category.icon !== undefined) updates.Icon = category.icon;
      if (category.sortOrder !== undefined) updates.SortOrder = category.sortOrder;
      if (category.parentCategoryId !== undefined) updates.ParentCategoryId = category.parentCategoryId;

      await this.sp.web.lists.getByTitle(this.SKILL_CATEGORIES_LIST).items
        .getById(id)
        .update(updates);

      return await this.getSkillCategoryById(id);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error updating skill category ${id}`, error);
      throw error;
    }
  }

  /**
   * Delete a skill category
   */
  public async deleteSkillCategory(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.SKILL_CATEGORIES_LIST).items
        .getById(id)
        .delete();
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error deleting skill category ${id}`, error);
      throw error;
    }
  }

  // ===================================================================
  // USER SKILLS
  // ===================================================================

  /**
   * Get all skills for a user
   */
  public async getUserSkills(userEmail: string): Promise<IUserSkill[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .select('Id', 'Title', 'UserId', 'UserEmail', 'SkillId', 'SkillCode', 'SelfRating', 'ManagerRating', 'VerifiedRating', 'LastAssessedDate', 'AssessedBy', 'Evidence', 'SkillSource', 'Notes')
        .filter(`UserEmail eq '${userEmail}'`)
        .top(500)() as IUserSkillListItem[];

      return items.map(this.mapUserSkillListItemToUserSkill);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching user skills for ${userEmail}`, error);
      throw error;
    }
  }

  /**
   * Get user skill by ID
   */
  public async getUserSkillById(id: number): Promise<IUserSkill> {
    try {
      const item = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .getById(id)
        .select('Id', 'Title', 'UserId', 'UserEmail', 'SkillId', 'SkillCode', 'SelfRating', 'ManagerRating', 'VerifiedRating', 'LastAssessedDate', 'AssessedBy', 'Evidence', 'SkillSource', 'Notes')() as IUserSkillListItem;

      return this.mapUserSkillListItemToUserSkill(item);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching user skill ${id}`, error);
      throw error;
    }
  }

  /**
   * Add a skill to user's profile
   */
  public async addUserSkill(userSkill: Partial<IUserSkillListItem>): Promise<IUserSkill> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items.add({
        Title: userSkill.Title,
        UserId: userSkill.UserId,
        UserEmail: userSkill.UserEmail,
        SkillId: userSkill.SkillId,
        SkillCode: userSkill.SkillCode,
        SelfRating: userSkill.SelfRating,
        SkillSource: userSkill.SkillSource || 'Self-Assessment',
        Notes: userSkill.Notes
      });

      return await this.getUserSkillById(result.data.Id);
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error adding user skill', error);
      throw error;
    }
  }

  /**
   * Update user skill rating
   */
  public async updateUserSkillRating(id: number, selfRating?: number, managerRating?: number): Promise<void> {
    try {
      const updates: Partial<IUserSkillListItem> = {
        LastAssessedDate: new Date().toISOString()
      };

      if (selfRating !== undefined) {
        updates.SelfRating = selfRating;
      }
      if (managerRating !== undefined) {
        updates.ManagerRating = managerRating;
      }

      await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error updating user skill ${id}`, error);
      throw error;
    }
  }

  /**
   * Remove skill from user's profile
   */
  public async removeUserSkill(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .getById(id)
        .delete();
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error removing user skill ${id}`, error);
      throw error;
    }
  }

  // ===================================================================
  // ROLE COMPETENCIES
  // ===================================================================

  /**
   * Get all role competencies
   */
  public async getRoleCompetencies(filter?: string): Promise<IRoleCompetency[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.ROLE_COMPETENCIES_LIST).items
        .select('Id', 'Title', 'RoleCode', 'Level', 'Department', 'RoleStatus', 'RequiredSkills', 'PreferredSkills', 'SuccessionPath')
        .orderBy('Title', true)
        .top(500);

      if (filter) {
        query = query.filter(filter);
      }

      const items = await query() as IRoleCompetencyListItem[];
      return items.map(this.mapRoleCompetencyListItemToRoleCompetency);
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error fetching role competencies', error);
      throw error;
    }
  }

  /**
   * Get active role competencies only
   */
  public async getActiveRoleCompetencies(): Promise<IRoleCompetency[]> {
    return this.getRoleCompetencies("RoleStatus eq 'Active'");
  }

  /**
   * Get role competency by ID
   */
  public async getRoleCompetencyById(id: number): Promise<IRoleCompetency> {
    try {
      const item = await this.sp.web.lists.getByTitle(this.ROLE_COMPETENCIES_LIST).items
        .getById(id)
        .select('Id', 'Title', 'RoleCode', 'Level', 'Department', 'RoleStatus', 'RequiredSkills', 'PreferredSkills', 'SuccessionPath')() as IRoleCompetencyListItem;

      return this.mapRoleCompetencyListItemToRoleCompetency(item);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching role competency ${id}`, error);
      throw error;
    }
  }

  /**
   * Get role competencies by department
   */
  public async getRoleCompetenciesByDepartment(department: string): Promise<IRoleCompetency[]> {
    return this.getRoleCompetencies(`Department eq '${department}'`);
  }

  // ===================================================================
  // SKILL TESTS
  // ===================================================================

  /**
   * Get all skill tests
   */
  public async getSkillTests(filter?: string): Promise<ISkillTestListItem[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.SKILL_TESTS_LIST).items
        .select('Id', 'Title', 'SkillId', 'SkillCode', 'Description', 'Questions', 'PassingScore', 'TimeLimit', 'DifficultyLevel', 'TestStatus')
        .orderBy('Title', true)
        .top(500);

      if (filter) {
        query = query.filter(filter);
      }

      const items = await query();
      return items as ISkillTestListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error fetching skill tests', error);
      throw error;
    }
  }

  /**
   * Get published skill tests
   */
  public async getPublishedSkillTests(): Promise<ISkillTestListItem[]> {
    return this.getSkillTests("TestStatus eq 'Published'");
  }

  /**
   * Get skill test by ID
   */
  public async getSkillTestById(id: number): Promise<ISkillTestListItem> {
    try {
      const item = await this.sp.web.lists.getByTitle(this.SKILL_TESTS_LIST).items
        .getById(id)
        .select('Id', 'Title', 'SkillId', 'SkillCode', 'Description', 'Questions', 'PassingScore', 'TimeLimit', 'DifficultyLevel', 'TestStatus')();

      return item as ISkillTestListItem;
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching skill test ${id}`, error);
      throw error;
    }
  }

  /**
   * Get tests for a specific skill
   */
  public async getTestsForSkill(skillId: number): Promise<ISkillTestListItem[]> {
    return this.getSkillTests(`SkillId eq ${skillId}`);
  }

  // ===================================================================
  // SKILL TEST ATTEMPTS
  // ===================================================================

  /**
   * Get test attempts for a user
   */
  public async getUserTestAttempts(userEmail: string): Promise<ISkillTestAttemptListItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .select('Id', 'Title', 'TestId', 'UserId', 'UserEmail', 'AttemptNumber', 'Score', 'Passed', 'StartTime', 'EndTime', 'Answers')
        .filter(`UserEmail eq '${userEmail}'`)
        .orderBy('StartTime', false)
        .top(100)();

      return items as ISkillTestAttemptListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching test attempts for ${userEmail}`, error);
      throw error;
    }
  }

  /**
   * Record a test attempt
   */
  public async recordTestAttempt(attempt: Partial<ISkillTestAttemptListItem>): Promise<ISkillTestAttemptListItem> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items.add({
        Title: attempt.Title,
        TestId: attempt.TestId,
        UserId: attempt.UserId,
        UserEmail: attempt.UserEmail,
        AttemptNumber: attempt.AttemptNumber || 1,
        Score: attempt.Score,
        Passed: attempt.Passed,
        StartTime: attempt.StartTime,
        EndTime: attempt.EndTime,
        Answers: attempt.Answers
      });

      return await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .getById(result.data.Id)
        .select('Id', 'Title', 'TestId', 'UserId', 'UserEmail', 'AttemptNumber', 'Score', 'Passed', 'StartTime', 'EndTime', 'Answers')() as ISkillTestAttemptListItem;
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error recording test attempt', error);
      throw error;
    }
  }

  /**
   * Start a new test attempt
   */
  public async startTestAttempt(testId: number, userId: number, userEmail: string): Promise<ISkillTestAttemptListItem> {
    try {
      // Get test details
      const test = await this.getSkillTestById(testId);

      // Count existing attempts
      const existingAttempts = await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .filter(`TestId eq ${testId} and UserEmail eq '${userEmail}'`)
        .select('Id')();

      const attemptNumber = existingAttempts.length + 1;

      // Create new attempt
      const result = await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items.add({
        Title: `${test.Title} - Attempt ${attemptNumber}`,
        TestId: testId,
        UserId: userId,
        UserEmail: userEmail,
        AttemptNumber: attemptNumber,
        StartTime: new Date().toISOString(),
        Score: 0,
        Passed: false
      });

      logger.info('SkillsCompetenciesService', `Test attempt ${attemptNumber} started for test ${testId}`);

      return await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .getById(result.data.Id)
        .select('Id', 'Title', 'TestId', 'UserId', 'UserEmail', 'AttemptNumber', 'Score', 'Passed', 'StartTime', 'EndTime', 'Answers')() as ISkillTestAttemptListItem;
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error starting test attempt', error);
      throw error;
    }
  }

  /**
   * Complete a test attempt with answers and score
   */
  public async completeTestAttempt(
    attemptId: number,
    answers: string,
    score: number,
    passed: boolean
  ): Promise<ISkillTestAttemptListItem> {
    try {
      await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .getById(attemptId)
        .update({
          EndTime: new Date().toISOString(),
          Answers: answers,
          Score: score,
          Passed: passed
        });

      const attempt = await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .getById(attemptId)
        .select('Id', 'Title', 'TestId', 'UserId', 'UserEmail', 'AttemptNumber', 'Score', 'Passed', 'StartTime', 'EndTime', 'Answers')() as ISkillTestAttemptListItem;

      logger.info('SkillsCompetenciesService', `Test attempt ${attemptId} completed with score ${score}, passed: ${passed}`);

      return attempt;
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error completing test attempt ${attemptId}`, error);
      throw error;
    }
  }

  /**
   * Get attempts for a specific test by user
   */
  public async getTestAttemptsByUser(testId: number, userEmail: string): Promise<ISkillTestAttemptListItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SKILL_TEST_ATTEMPTS_LIST).items
        .select('Id', 'Title', 'TestId', 'UserId', 'UserEmail', 'AttemptNumber', 'Score', 'Passed', 'StartTime', 'EndTime', 'Answers')
        .filter(`TestId eq ${testId} and UserEmail eq '${userEmail}'`)
        .orderBy('AttemptNumber', false)();

      return items as ISkillTestAttemptListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching test attempts for test ${testId}`, error);
      throw error;
    }
  }

  /**
   * Update user skill level when they pass a test
   */
  public async updateSkillLevelOnTestPass(userEmail: string, skillId: number, validatedLevel: number): Promise<void> {
    try {
      // Find user's skill record
      const userSkills = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .filter(`UserEmail eq '${userEmail}' and SkillId eq ${skillId}`)
        .select('Id', 'VerifiedRating')
        .top(1)();

      if (userSkills.length > 0) {
        // Only update if the new level is higher
        const currentVerified = userSkills[0].VerifiedRating || 0;
        if (validatedLevel > currentVerified) {
          await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
            .getById(userSkills[0].Id)
            .update({
              VerifiedRating: validatedLevel,
              LastAssessedDate: new Date().toISOString(),
              AssessedBy: 'Skill Test',
              SkillSource: 'Skill Assessment'
            });
          logger.info('SkillsCompetenciesService', `Updated skill ${skillId} to verified level ${validatedLevel} for ${userEmail}`);
        }
      } else {
        // Create new user skill record
        const skill = await this.getSkillById(skillId);
        await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items.add({
          Title: skill.Title,
          UserEmail: userEmail,
          SkillId: skillId,
          SkillCode: skill.SkillCode,
          VerifiedRating: validatedLevel,
          LastAssessedDate: new Date().toISOString(),
          AssessedBy: 'Skill Test',
          SkillSource: 'Skill Assessment'
        });
        logger.info('SkillsCompetenciesService', `Created new skill record for ${userEmail} with verified level ${validatedLevel}`);
      }
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error updating skill level for ${userEmail}`, error);
      throw error;
    }
  }

  /**
   * Check if user can attempt a test (based on max attempts)
   */
  public async canAttemptTest(testId: number, userEmail: string): Promise<{ canAttempt: boolean; remainingAttempts: number; message: string }> {
    try {
      const test = await this.getSkillTestById(testId);
      const attempts = await this.getTestAttemptsByUser(testId, userEmail);

      // Check if user has passed already
      const passedAttempt = attempts.find(a => a.Passed);
      if (passedAttempt) {
        return {
          canAttempt: false,
          remainingAttempts: 0,
          message: 'You have already passed this test.'
        };
      }

      // Check max attempts (assuming MaxAttempts is stored, defaulting to 3)
      const maxAttempts = 3; // Could be stored in test config
      const remaining = maxAttempts - attempts.length;

      if (remaining <= 0) {
        return {
          canAttempt: false,
          remainingAttempts: 0,
          message: 'You have used all available attempts for this test.'
        };
      }

      return {
        canAttempt: true,
        remainingAttempts: remaining,
        message: `You have ${remaining} attempt${remaining !== 1 ? 's' : ''} remaining.`
      };
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error checking test eligibility`, error);
      throw error;
    }
  }

  // ===================================================================
  // SKILLS ASSESSMENTS
  // ===================================================================

  /**
   * Get assessments for a user
   */
  public async getUserAssessments(userEmail: string): Promise<ISkillsAssessmentListItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .select('Id', 'Title', 'UserId', 'UserEmail', 'RoleId', 'AssessmentDate', 'OverallScore', 'AssessmentStatus', 'AssessorId', 'AssessorName', 'Comments')
        .filter(`UserEmail eq '${userEmail}'`)
        .orderBy('AssessmentDate', false)
        .top(100)();

      return items as ISkillsAssessmentListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching assessments for ${userEmail}`, error);
      throw error;
    }
  }

  /**
   * Create a new assessment
   */
  public async createAssessment(assessment: Partial<ISkillsAssessmentListItem>): Promise<ISkillsAssessmentListItem> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items.add({
        Title: assessment.Title,
        UserId: assessment.UserId,
        UserEmail: assessment.UserEmail,
        RoleId: assessment.RoleId,
        AssessmentDate: assessment.AssessmentDate || new Date().toISOString(),
        OverallScore: assessment.OverallScore,
        AssessmentStatus: assessment.AssessmentStatus || 'Draft',
        AssessorId: assessment.AssessorId,
        AssessorName: assessment.AssessorName,
        Comments: assessment.Comments
      });

      return await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .getById(result.data.Id)() as ISkillsAssessmentListItem;
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error creating assessment', error);
      throw error;
    }
  }

  /**
   * Get pending assessments awaiting verification
   */
  public async getPendingAssessments(): Promise<ISkillsAssessmentListItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .select('Id', 'Title', 'UserId', 'UserEmail', 'RoleId', 'AssessmentDate', 'OverallScore', 'AssessmentStatus', 'AssessorId', 'AssessorName', 'Comments')
        .filter("AssessmentStatus eq 'Pending Verification'")
        .orderBy('AssessmentDate', true)
        .top(100)();

      return items as ISkillsAssessmentListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error fetching pending assessments', error);
      throw error;
    }
  }

  /**
   * Get assessment by ID
   */
  public async getAssessmentById(id: number): Promise<ISkillsAssessmentListItem> {
    try {
      const item = await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .getById(id)
        .select('Id', 'Title', 'UserId', 'UserEmail', 'RoleId', 'AssessmentDate', 'OverallScore', 'AssessmentStatus', 'AssessorId', 'AssessorName', 'Comments')();

      return item as ISkillsAssessmentListItem;
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching assessment ${id}`, error);
      throw error;
    }
  }

  /**
   * Update an existing assessment
   */
  public async updateAssessment(id: number, assessment: {
    title?: string;
    overallScore?: number;
    status?: string;
    assessorName?: string;
    comments?: string;
  }): Promise<ISkillsAssessmentListItem> {
    try {
      const updates: Record<string, unknown> = {};

      if (assessment.title !== undefined) updates.Title = assessment.title;
      if (assessment.overallScore !== undefined) updates.OverallScore = assessment.overallScore;
      if (assessment.status !== undefined) updates.AssessmentStatus = assessment.status;
      if (assessment.assessorName !== undefined) updates.AssessorName = assessment.assessorName;
      if (assessment.comments !== undefined) updates.Comments = assessment.comments;

      await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .getById(id)
        .update(updates);

      return await this.getAssessmentById(id);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error updating assessment ${id}`, error);
      throw error;
    }
  }

  /**
   * Approve a pending assessment and update the user's verified skill level
   */
  public async approveAssessment(assessmentId: number, verifiedLevel: number, approverName: string, comments?: string): Promise<void> {
    try {
      // Update assessment status
      await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .getById(assessmentId)
        .update({
          AssessmentStatus: 'Approved',
          OverallScore: verifiedLevel,
          AssessorName: approverName,
          Comments: comments || ''
        });

      // Get assessment details to update user skill
      const assessment = await this.getAssessmentById(assessmentId);

      // Find and update user skill with verified level
      const userSkills = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .filter(`UserEmail eq '${assessment.UserEmail}'`)
        .top(500)();

      // Update the first skill found (assumes Title contains skill info)
      // In a real implementation, would need skill ID association
      if (userSkills.length > 0) {
        await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
          .getById(userSkills[0].Id)
          .update({
            VerifiedRating: verifiedLevel,
            AssessedBy: approverName,
            LastAssessedDate: new Date().toISOString()
          });
      }

      logger.info('SkillsCompetenciesService', `Assessment ${assessmentId} approved with level ${verifiedLevel}`);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error approving assessment ${assessmentId}`, error);
      throw error;
    }
  }

  /**
   * Reject a pending assessment
   */
  public async rejectAssessment(assessmentId: number, rejectorName: string, reason: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .getById(assessmentId)
        .update({
          AssessmentStatus: 'Rejected',
          AssessorName: rejectorName,
          Comments: `Rejection reason: ${reason}`
        });

      logger.info('SkillsCompetenciesService', `Assessment ${assessmentId} rejected`);
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error rejecting assessment ${assessmentId}`, error);
      throw error;
    }
  }

  /**
   * Delete an assessment
   */
  public async deleteAssessment(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .getById(id)
        .delete();
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error deleting assessment ${id}`, error);
      throw error;
    }
  }

  // ===================================================================
  // MANAGER ASSESSMENT WORKFLOW
  // ===================================================================

  /**
   * Submit a manager skill rating for a team member
   * Updates the user skill record with manager's rating
   */
  public async submitManagerSkillRating(
    userEmail: string,
    skillId: number,
    managerRating: number,
    managerName: string,
    comments?: string
  ): Promise<void> {
    try {
      // Find existing user skill record
      const existingSkills = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
        .filter(`UserEmail eq '${userEmail}' and SkillId eq ${skillId}`)
        .select('Id')
        .top(1)();

      if (existingSkills.length > 0) {
        // Update existing record
        await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
          .getById(existingSkills[0].Id)
          .update({
            ManagerRating: managerRating,
            AssessedBy: managerName,
            LastAssessedDate: new Date().toISOString(),
            Notes: comments || ''
          });
        logger.info('SkillsCompetenciesService', `Manager rating updated for skill ${skillId}, user ${userEmail}`);
      } else {
        // Create new user skill record
        const skill = await this.getSkillById(skillId);
        await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items.add({
          Title: skill.Title,
          UserEmail: userEmail,
          SkillId: skillId,
          SkillCode: skill.SkillCode,
          ManagerRating: managerRating,
          AssessedBy: managerName,
          LastAssessedDate: new Date().toISOString(),
          SkillSource: 'Manager Assessment',
          Notes: comments || ''
        });
        logger.info('SkillsCompetenciesService', `Manager rating created for skill ${skillId}, user ${userEmail}`);
      }
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error submitting manager skill rating`, error);
      throw error;
    }
  }

  /**
   * Create a formal manager assessment for a team member
   * This creates an assessment record and updates all skill ratings
   */
  public async createManagerAssessment(
    targetUserEmail: string,
    targetUserId: number,
    managerName: string,
    managerId: number,
    skillRatings: Array<{ skillId: number; rating: number; comments?: string }>,
    overallComments?: string
  ): Promise<ISkillsAssessmentListItem> {
    try {
      // Calculate overall score as average of all ratings
      const overallScore = Math.round(
        skillRatings.reduce((sum, sr) => sum + sr.rating, 0) / skillRatings.length
      );

      // Create the assessment record
      const assessment = await this.createAssessment({
        Title: `Manager Assessment - ${new Date().toLocaleDateString()}`,
        UserId: targetUserId,
        UserEmail: targetUserEmail,
        AssessmentDate: new Date().toISOString(),
        OverallScore: overallScore,
        AssessmentStatus: 'Pending Verification',
        AssessorId: managerId,
        AssessorName: managerName,
        Comments: overallComments
      });

      // Update individual skill ratings
      for (const skillRating of skillRatings) {
        await this.submitManagerSkillRating(
          targetUserEmail,
          skillRating.skillId,
          skillRating.rating,
          managerName,
          skillRating.comments
        );
      }

      logger.info('SkillsCompetenciesService', `Manager assessment created for ${targetUserEmail}`);
      return assessment;
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error creating manager assessment', error);
      throw error;
    }
  }

  /**
   * Get assessments submitted by a manager
   */
  public async getManagerSubmittedAssessments(managerId: number): Promise<ISkillsAssessmentListItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SKILLS_ASSESSMENTS_LIST).items
        .select('Id', 'Title', 'UserId', 'UserEmail', 'RoleId', 'AssessmentDate', 'OverallScore', 'AssessmentStatus', 'AssessorId', 'AssessorName', 'Comments')
        .filter(`AssessorId eq ${managerId}`)
        .orderBy('AssessmentDate', false)
        .top(100)();

      return items as ISkillsAssessmentListItem[];
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error fetching manager assessments for ${managerId}`, error);
      throw error;
    }
  }

  // ===================================================================
  // GAP ANALYSIS
  // ===================================================================

  /**
   * Calculate skills gap analysis for a user against a role
   */
  public async calculateGapAnalysis(userEmail: string, roleId: number): Promise<ISkillsGapAnalysis> {
    try {
      // Get user skills
      const userSkills = await this.getUserSkills(userEmail);

      // Get role competencies
      const role = await this.getRoleCompetencyById(roleId);

      // Get all skills for mapping
      const allSkills = await this.getActiveSkills();
      const skillsMap = new Map(allSkills.map(s => [s.Id, s]));

      const gaps: ISkillGap[] = [];
      const strengths: ISkillStrength[] = [];

      // Create a map of user's skill levels
      const userSkillLevels = new Map<number, number>();
      userSkills.forEach(us => {
        const level = us.VerifiedLevel || us.ManagerRating || us.SelfRating || 0;
        userSkillLevels.set(us.SkillId, level as number);
      });

      // Analyze required skills
      for (const req of role.RequiredSkills) {
        const currentLevel = userSkillLevels.get(req.skillId) || 0;
        const requiredLevel = req.requiredLevel as number;
        const gap = requiredLevel - currentLevel;
        const skill = skillsMap.get(req.skillId);

        if (gap > 0 && skill) {
          gaps.push({
            skill: skill,
            skillId: req.skillId,
            skillName: req.skillName,
            requiredLevel: req.requiredLevel,
            currentLevel: currentLevel as ProficiencyLevel,
            gap: gap,
            priority: gap >= 2 ? 'Critical' : gap >= 1 ? 'High' : 'Medium',
            suggestedCourses: [], // Would need training catalog integration
            estimatedTimeToClose: gap * 10 // Rough estimate: 10 hours per level
          });
        } else if (gap < 0 && skill) {
          strengths.push({
            skill: skill,
            skillId: req.skillId,
            skillName: req.skillName,
            currentLevel: currentLevel as ProficiencyLevel,
            requiredLevel: req.requiredLevel,
            exceeds: Math.abs(gap)
          });
        }
      }

      // Calculate overall readiness
      const totalRequired = role.RequiredSkills.length;
      const metOrExceeded = totalRequired - gaps.length;
      const overallReadiness = totalRequired > 0 ? Math.round((metOrExceeded / totalRequired) * 100) : 100;

      // Calculate estimated time to close all gaps
      const totalTimeToClose = gaps.reduce((sum, g) => sum + g.estimatedTimeToClose, 0);

      return {
        userId: 0, // Would need to be looked up
        userName: userEmail,
        roleId: roleId,
        roleName: role.RoleTitle,
        analysisDate: new Date(),
        overallReadiness: overallReadiness,
        gaps: gaps.sort((a, b) => {
          const priorityOrder = { 'Critical': 0, 'High': 1, 'Medium': 2, 'Low': 3 };
          return priorityOrder[a.priority] - priorityOrder[b.priority];
        }),
        strengths: strengths,
        recommendedTraining: [],
        recommendedCertifications: [],
        estimatedTimeToClose: totalTimeToClose,
        priorityActions: gaps.slice(0, 3).map(g => `Close ${g.skillName} gap (${g.gap} levels needed)`)
      };
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error calculating gap analysis', error);
      throw error;
    }
  }

  // ===================================================================
  // TRAINING COMPLETION INTEGRATION
  // ===================================================================

  /**
   * Update user skills when they complete a training course
   * Maps course difficulty to skill proficiency level
   */
  public async updateSkillsOnCourseCompletion(
    userEmail: string,
    courseId: number,
    skillIds: number[],
    difficultyLevel: string
  ): Promise<void> {
    try {
      // Map difficulty level to proficiency
      const proficiencyFromDifficulty = this.mapDifficultyToProficiency(difficultyLevel);

      // Update each skill associated with the course
      for (const skillId of skillIds) {
        // Check if user already has this skill at a higher level
        const existingSkills = await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
          .filter(`UserEmail eq '${userEmail}' and SkillId eq ${skillId}`)
          .select('Id', 'VerifiedRating', 'ManagerRating', 'SelfRating')
          .top(1)();

        if (existingSkills.length > 0) {
          const existing = existingSkills[0];
          const currentLevel = existing.VerifiedRating || existing.ManagerRating || existing.SelfRating || 0;

          // Only update if the training grants a higher level
          if (proficiencyFromDifficulty > currentLevel) {
            await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items
              .getById(existing.Id)
              .update({
                VerifiedRating: proficiencyFromDifficulty,
                LastAssessedDate: new Date().toISOString(),
                AssessedBy: 'Training Completion',
                SkillSource: 'Training Completion'
              });
            logger.info('SkillsCompetenciesService', `Skill ${skillId} updated to level ${proficiencyFromDifficulty} for ${userEmail} from course ${courseId}`);
          }
        } else {
          // Create new user skill record from training
          const skill = await this.getSkillById(skillId);
          await this.sp.web.lists.getByTitle(this.USER_SKILLS_LIST).items.add({
            Title: skill.Title,
            UserEmail: userEmail,
            SkillId: skillId,
            SkillCode: skill.SkillCode,
            VerifiedRating: proficiencyFromDifficulty,
            LastAssessedDate: new Date().toISOString(),
            AssessedBy: 'Training Completion',
            SkillSource: 'Training Completion',
            Notes: `Acquired through course ID: ${courseId}`
          });
          logger.info('SkillsCompetenciesService', `New skill ${skillId} at level ${proficiencyFromDifficulty} created for ${userEmail} from course ${courseId}`);
        }
      }
    } catch (error) {
      logger.error('SkillsCompetenciesService', `Error updating skills on course completion for ${userEmail}`, error);
      throw error;
    }
  }

  /**
   * Map course difficulty level to skill proficiency
   */
  private mapDifficultyToProficiency(difficulty: string): number {
    switch (difficulty.toLowerCase()) {
      case 'beginner':
      case 'foundational':
        return 2; // Beginner proficiency
      case 'intermediate':
        return 3; // Competent proficiency
      case 'advanced':
        return 4; // Proficient
      case 'expert':
        return 5; // Expert
      default:
        return 2; // Default to beginner
    }
  }

  /**
   * Record course completion and update associated skills
   * This is the main entry point called when a course is completed
   */
  public async recordCourseCompletionWithSkillUpdate(
    userEmail: string,
    courseId: number,
    courseTitle: string,
    skillIds: number[],
    difficultyLevel: string,
    score?: number
  ): Promise<void> {
    try {
      logger.info('SkillsCompetenciesService', `Recording course completion for ${userEmail}: ${courseTitle}`);

      // Update skills if course has associated skills
      if (skillIds && skillIds.length > 0) {
        await this.updateSkillsOnCourseCompletion(userEmail, courseId, skillIds, difficultyLevel);
      }

      // Record completion in assessment log (optional tracking)
      await this.createAssessment({
        Title: `Course Completion: ${courseTitle}`,
        UserEmail: userEmail,
        AssessmentDate: new Date().toISOString(),
        OverallScore: score || 100,
        AssessmentStatus: 'Approved',
        AssessorName: 'System',
        Comments: `Completed course ${courseTitle} (ID: ${courseId}) at ${difficultyLevel} level`
      });

      logger.info('SkillsCompetenciesService', `Course completion recorded successfully for ${userEmail}`);
    } catch (error) {
      logger.error('SkillsCompetenciesService', 'Error recording course completion', error);
      throw error;
    }
  }

  // ===================================================================
  // MAPPING FUNCTIONS
  // ===================================================================

  private mapSkillListItemToSkill(item: ISkillListItem): ISkill {
    return {
      Id: item.Id,
      Title: item.Title,
      SkillCode: item.SkillCode,
      Description: item.Description || '',
      Domain: item.Domain as SkillDomain,
      CategoryId: item.SkillCategoryId,
      ProficiencyLevels: [], // Would need separate lookup
      IsCore: item.IsCore,
      IsActive: item.SkillStatus === 'Active'
    };
  }

  private mapUserSkillListItemToUserSkill(item: IUserSkillListItem): IUserSkill {
    return {
      Id: item.Id,
      Title: item.Title,
      UserId: item.UserId,
      UserEmail: item.UserEmail,
      SkillId: item.SkillId,
      SelfRating: item.SelfRating as ProficiencyLevel,
      ManagerRating: item.ManagerRating as ProficiencyLevel,
      VerifiedLevel: item.VerifiedRating as ProficiencyLevel,
      LastAssessedDate: item.LastAssessedDate ? new Date(item.LastAssessedDate) : undefined,
      AssessedByName: item.AssessedBy,
      Evidence: item.Evidence ? JSON.parse(item.Evidence) : undefined,
      Source: item.SkillSource as SkillSource,
      Notes: item.Notes,
      EffectiveLevel: (item.VerifiedRating || item.ManagerRating || item.SelfRating) as ProficiencyLevel
    };
  }

  private mapRoleCompetencyListItemToRoleCompetency(item: IRoleCompetencyListItem): IRoleCompetency {
    let requiredSkills: IRoleSkillRequirement[] = [];
    let preferredSkills: IRoleSkillRequirement[] = [];
    let successionPath: ICareerStep[] = [];

    try {
      if (item.RequiredSkills) {
        requiredSkills = JSON.parse(item.RequiredSkills);
      }
    } catch {
      logger.warn('SkillsCompetenciesService', `Failed to parse RequiredSkills for role ${item.Id}`);
    }

    try {
      if (item.PreferredSkills) {
        preferredSkills = JSON.parse(item.PreferredSkills);
      }
    } catch {
      logger.warn('SkillsCompetenciesService', `Failed to parse PreferredSkills for role ${item.Id}`);
    }

    try {
      if (item.SuccessionPath) {
        successionPath = JSON.parse(item.SuccessionPath);
      }
    } catch {
      logger.warn('SkillsCompetenciesService', `Failed to parse SuccessionPath for role ${item.Id}`);
    }

    return {
      Id: item.Id,
      Title: item.Title,
      RoleTitle: item.Title,
      RoleFamily: item.RoleCode,
      Department: item.Department,
      Level: item.Level,
      RequiredSkills: requiredSkills,
      PreferredSkills: preferredSkills,
      SuccessionPath: successionPath,
      IsActive: item.RoleStatus === 'Active',
      EffectiveDate: new Date(),
      Version: 1
    };
  }
}
