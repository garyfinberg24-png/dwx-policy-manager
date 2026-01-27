// @ts-nocheck
/**
 * TaskLibraryService
 * Service for loading and managing task library data from JSON files
 * Provides access to tasks, templates, and categories for the Process Wizard
 */

import tasksData from '../data/taskLibrary/tasks.json';
import templatesData from '../data/taskLibrary/templates.json';
import categoriesData from '../data/taskLibrary/categories.json';
import {
  ITaskLibraryItem,
  IProcessTemplate,
  ITaskCategoryDefinition,
  TaskPhase,
  TemplateTier,
  ISelectedTask,
  IProcessConfiguration
} from '../models/ITaskLibrary';
import { ProcessType, TaskCategory } from '../models/ICommon';

export class TaskLibraryService {
  private static instance: TaskLibraryService;
  private tasks: ITaskLibraryItem[];
  private templates: IProcessTemplate[];
  private categories: ITaskCategoryDefinition[];

  private constructor() {
    this.tasks = tasksData.tasks as ITaskLibraryItem[];
    this.templates = templatesData.templates as IProcessTemplate[];
    this.categories = categoriesData.categories as ITaskCategoryDefinition[];
  }

  /**
   * Get singleton instance
   */
  public static getInstance(): TaskLibraryService {
    if (!TaskLibraryService.instance) {
      TaskLibraryService.instance = new TaskLibraryService();
    }
    return TaskLibraryService.instance;
  }

  // ============================================================================
  // TASK OPERATIONS
  // ============================================================================

  /**
   * Get all tasks from library
   */
  public getAllTasks(): ITaskLibraryItem[] {
    return [...this.tasks];
  }

  /**
   * Get task by ID
   */
  public getTaskById(taskId: string): ITaskLibraryItem | undefined {
    return this.tasks.find(t => t.id === taskId);
  }

  /**
   * Get tasks by phase
   */
  public getTasksByPhase(phase: TaskPhase): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.phase === phase);
  }

  /**
   * Get tasks by category
   */
  public getTasksByCategory(category: TaskCategory | string): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.category === category);
  }

  /**
   * Get tasks by department
   */
  public getTasksByDepartment(department: string): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.defaultDepartment === department);
  }

  /**
   * Get tasks by tier
   */
  public getTasksByTier(tier: TemplateTier): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.includedInTiers.includes(tier));
  }

  /**
   * Get tasks by process type
   */
  public getTasksByProcessType(processType: ProcessType): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.applicableProcessTypes.includes(processType));
  }

  /**
   * Get common tasks (frequently used)
   */
  public getCommonTasks(): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.isCommon);
  }

  /**
   * Get optional tasks
   */
  public getOptionalTasks(): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.isOptional);
  }

  /**
   * Search tasks by keyword
   */
  public searchTasks(keyword: string): ITaskLibraryItem[] {
    const lowerKeyword = keyword.toLowerCase();
    return this.tasks.filter(t =>
      t.title.toLowerCase().includes(lowerKeyword) ||
      t.description.toLowerCase().includes(lowerKeyword) ||
      t.tags.some(tag => tag.toLowerCase().includes(lowerKeyword))
    );
  }

  /**
   * Get tasks with automation
   */
  public getAutomatedTasks(): ITaskLibraryItem[] {
    return this.tasks.filter(t =>
      t.automationRules && t.automationRules.length > 0
    );
  }

  /**
   * Get critical tasks
   */
  public getCriticalTasks(): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.priority === 'Critical');
  }

  /**
   * Get tasks that block other tasks
   */
  public getBlockingTasks(): ITaskLibraryItem[] {
    return this.tasks.filter(t => t.blocksOtherTasks);
  }

  /**
   * Get task dependencies
   */
  public getTaskDependencies(taskId: string): ITaskLibraryItem[] {
    const task = this.getTaskById(taskId);
    if (!task || !task.dependencies || task.dependencies.length === 0) {
      return [];
    }
    return task.dependencies
      .map(depId => this.getTaskById(depId))
      .filter(t => t !== undefined) as ITaskLibraryItem[];
  }

  /**
   * Get tasks that depend on this task
   */
  public getDependentTasks(taskId: string): ITaskLibraryItem[] {
    return this.tasks.filter(t =>
      t.dependencies && t.dependencies.includes(taskId)
    );
  }

  // ============================================================================
  // TEMPLATE OPERATIONS
  // ============================================================================

  /**
   * Get all templates
   */
  public getAllTemplates(): IProcessTemplate[] {
    return [...this.templates];
  }

  /**
   * Get template by ID
   */
  public getTemplateById(templateId: string): IProcessTemplate | undefined {
    return this.templates.find(t => t.id === templateId);
  }

  /**
   * Get templates by tier
   */
  public getTemplatesByTier(tier: TemplateTier): IProcessTemplate[] {
    return this.templates.filter(t => t.tier === tier);
  }

  /**
   * Get templates by process type
   */
  public getTemplatesByProcessType(processType: ProcessType): IProcessTemplate[] {
    return this.templates.filter(t => t.processType === processType);
  }

  /**
   * Get recommended template based on criteria
   */
  public getRecommendedTemplate(criteria: {
    processType: ProcessType;
    companySize?: string;
    industry?: string;
    hasRemoteEmployees?: boolean;
  }): IProcessTemplate | undefined {
    let candidates = this.getTemplatesByProcessType(criteria.processType);

    // Filter by company size if provided
    if (criteria.companySize) {
      candidates = candidates.filter(t =>
        !t.companySize || t.companySize.includes(criteria.companySize!)
      );
    }

    // Filter by industry if provided
    if (criteria.industry) {
      candidates = candidates.filter(t =>
        !t.industries || t.industries.includes(criteria.industry!)
      );
    }

    // If remote employees, prefer templates with remote support
    if (criteria.hasRemoteEmployees) {
      const remoteTemplates = candidates.filter(t =>
        t.tasks.includes('task-remote-specialty-001') ||
        t.tasks.includes('task-remote-specialty-002')
      );
      if (remoteTemplates.length > 0) {
        candidates = remoteTemplates;
      }
    }

    // Return highest rated or most used
    if (candidates.length === 0) return undefined;
    return candidates.sort((a, b) => {
      const scoreA = (a.rating || 0) * 0.5 + (a.usageCount || 0) * 0.001;
      const scoreB = (b.rating || 0) * 0.5 + (b.usageCount || 0) * 0.001;
      return scoreB - scoreA;
    })[0];
  }

  /**
   * Get tasks for template
   */
  public getTasksForTemplate(templateId: string): ITaskLibraryItem[] {
    const template = this.getTemplateById(templateId);
    if (!template) return [];

    return template.tasks
      .map(taskId => this.getTaskById(taskId))
      .filter(t => t !== undefined) as ITaskLibraryItem[];
  }

  /**
   * Get template statistics
   */
  public getTemplateStats(templateId: string): {
    taskCount: number;
    phaseBreakdown: Record<string, number>;
    departmentBreakdown: Record<string, number>;
    automationPercentage: number;
    estimatedHours: number;
  } | null {
    const tasks = this.getTasksForTemplate(templateId);
    if (tasks.length === 0) return null;

    const phaseBreakdown: Record<string, number> = {};
    const departmentBreakdown: Record<string, number> = {};
    let automatedCount = 0;
    let totalHours = 0;

    tasks.forEach(task => {
      // Phase breakdown
      phaseBreakdown[task.phase] = (phaseBreakdown[task.phase] || 0) + 1;

      // Department breakdown
      departmentBreakdown[task.defaultDepartment] =
        (departmentBreakdown[task.defaultDepartment] || 0) + 1;

      // Automation count
      if (task.automationRules && task.automationRules.length > 0) {
        automatedCount++;
      }

      // Total hours
      totalHours += task.estimatedHours;
    });

    return {
      taskCount: tasks.length,
      phaseBreakdown,
      departmentBreakdown,
      automationPercentage: Math.round((automatedCount / tasks.length) * 100),
      estimatedHours: totalHours
    };
  }

  // ============================================================================
  // CATEGORY OPERATIONS
  // ============================================================================

  /**
   * Get all categories
   */
  public getAllCategories(): ITaskCategoryDefinition[] {
    return [...this.categories];
  }

  /**
   * Get category by name
   */
  public getCategoryByName(categoryName: string): ITaskCategoryDefinition | undefined {
    return this.categories.find(c => c.category === categoryName);
  }

  /**
   * Get categories by department
   */
  public getCategoriesByDepartment(department: string): ITaskCategoryDefinition[] {
    return this.categories.filter(c => c.department === department);
  }

  /**
   * Get category icon
   */
  public getCategoryIcon(categoryName: string): string {
    const category = this.getCategoryByName(categoryName);
    return category?.icon || 'More';
  }

  /**
   * Get category color
   */
  public getCategoryColor(categoryName: string): string {
    const category = this.getCategoryByName(categoryName);
    return category?.color || '#737373';
  }

  // ============================================================================
  // CONFIGURATION HELPERS
  // ============================================================================

  /**
   * Create process configuration from template
   */
  public createConfigurationFromTemplate(
    templateId: string,
    processDetails: {
      employeeName: string;
      employeeEmail?: string;
      startDate: Date;
      department: string;
      jobTitle: string;
      managerId?: number;
      buddyId?: number;
    },
    options?: {
      includeOptionalTasks?: boolean;
      customTaskIds?: string[];
      excludeTaskIds?: string[];
    }
  ): IProcessConfiguration | null {
    const template = this.getTemplateById(templateId);
    if (!template) return null;

    let taskIds = [...template.tasks];

    // Add custom tasks
    if (options?.customTaskIds) {
      taskIds = [...taskIds, ...options.customTaskIds];
    }

    // Remove excluded tasks
    if (options?.excludeTaskIds) {
      taskIds = taskIds.filter(id => !options.excludeTaskIds!.includes(id));
    }

    // Filter optional tasks if requested
    if (options?.includeOptionalTasks === false) {
      const allTasks = this.getAllTasks();
      taskIds = taskIds.filter(id => {
        const task = allTasks.find(t => t.id === id);
        return task && !task.isOptional;
      });
    }

    // Create selected tasks
    const selectedTasks: ISelectedTask[] = taskIds.map(taskId => ({
      libraryTaskId: taskId
    }));

    return {
      templateId: template.id,
      templateTier: template.tier,
      processType: template.processType,
      ...processDetails,
      tasks: selectedTasks,
      enableAutomation: true,
      notifyStakeholders: true,
      sendWelcomeEmail: template.processType === ProcessType.Joiner,
      scheduleKickoffMeeting: true,
      createdDate: new Date()
    };
  }

  /**
   * Validate task configuration
   */
  public validateConfiguration(config: IProcessConfiguration): {
    isValid: boolean;
    errors: string[];
    warnings: string[];
  } {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Check required fields
    if (!config.employeeName || config.employeeName.trim() === '') {
      errors.push('Employee name is required');
    }

    if (!config.department || config.department.trim() === '') {
      errors.push('Department is required');
    }

    if (!config.jobTitle || config.jobTitle.trim() === '') {
      errors.push('Job title is required');
    }

    if (!config.startDate) {
      errors.push('Start date is required');
    }

    if (!config.tasks || config.tasks.length === 0) {
      errors.push('At least one task must be selected');
    }

    // Check for missing dependencies
    const allTasks = this.getAllTasks();
    const selectedTaskIds = config.tasks.map(t => t.libraryTaskId);

    config.tasks.forEach(selectedTask => {
      const libraryTask = allTasks.find(t => t.id === selectedTask.libraryTaskId);
      if (libraryTask && libraryTask.dependencies) {
        const missingDeps = libraryTask.dependencies.filter(
          depId => !selectedTaskIds.includes(depId)
        );
        if (missingDeps.length > 0) {
          warnings.push(
            `Task "${libraryTask.title}" has missing dependencies: ${missingDeps.join(', ')}`
          );
        }
      }
    });

    // Check for blocking tasks
    const blockingTasks = config.tasks.filter(t => {
      const libraryTask = allTasks.find(lt => lt.id === t.libraryTaskId);
      return libraryTask && libraryTask.blocksOtherTasks;
    });

    if (blockingTasks.length > 3) {
      warnings.push(
        `Process has ${blockingTasks.length} blocking tasks which may create bottlenecks`
      );
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Calculate process timeline
   */
  public calculateTimeline(config: IProcessConfiguration): {
    startDate: Date;
    endDate: Date;
    totalDays: number;
    milestones: Array<{
      date: Date;
      phase: TaskPhase;
      taskCount: number;
    }>;
  } {
    const allTasks = this.getAllTasks();
    const selectedTasks = config.tasks
      .map(t => allTasks.find(lt => lt.id === t.libraryTaskId))
      .filter(t => t !== undefined) as ITaskLibraryItem[];

    // Calculate min and max offsets
    let minOffset = 0;
    let maxOffset = 0;

    selectedTasks.forEach(task => {
      const offset = task.defaultDaysOffset;
      if (offset < minOffset) minOffset = offset;
      if (offset > maxOffset) maxOffset = offset;
    });

    const startDate = new Date(config.startDate);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + (maxOffset - minOffset));

    // Group by phase for milestones
    const phaseGroups: Record<string, ITaskLibraryItem[]> = {};
    selectedTasks.forEach(task => {
      if (!phaseGroups[task.phase]) {
        phaseGroups[task.phase] = [];
      }
      phaseGroups[task.phase].push(task);
    });

    const milestones = Object.entries(phaseGroups).map(([phase, tasks]) => {
      // Calculate average offset for phase
      const avgOffset = tasks.reduce((sum, t) => sum + t.defaultDaysOffset, 0) / tasks.length;
      const milestoneDate = new Date(startDate);
      milestoneDate.setDate(milestoneDate.getDate() + Math.round(avgOffset));

      return {
        date: milestoneDate,
        phase: phase as TaskPhase,
        taskCount: tasks.length
      };
    });

    return {
      startDate,
      endDate,
      totalDays: maxOffset - minOffset,
      milestones: milestones.sort((a, b) => a.date.getTime() - b.date.getTime())
    };
  }
}

// Export singleton instance
export const taskLibraryService = TaskLibraryService.getInstance();
