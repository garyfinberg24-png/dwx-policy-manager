// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import { PolicyAnalyticsService } from './PolicyAnalyticsService';

/**
 * Integration service to bridge Policy Management with JML Processes
 * Links policy compliance requirements to onboarding/process workflows
 */
export class PolicyJMLIntegrationService {
  private sp: SPFI;
  private analyticsService: PolicyAnalyticsService;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.analyticsService = new PolicyAnalyticsService(sp);
  }

  /**
   * Create policy compliance tasks for a new JML process
   */
  public async createPolicyTasksForProcess(
    processId: number,
    employeeDepartment: string,
    jobTitle: string,
    employeeEmail: string,
    assignedToId: number,
    startDate: Date
  ): Promise<number[]> {
    try {
      // Get mandatory policies for this department/role
      const mandatoryPolicies = await this.getMandatoryPoliciesForRole(
        employeeDepartment,
        jobTitle
      );

      const taskIds: number[] = [];

      // Create task assignments for each mandatory policy
      for (const policy of mandatoryPolicies) {
        // 1. Create "Read Policy" task
        const readTaskId = await this.createPolicyTask(
          processId,
          assignedToId,
          employeeEmail,
          {
            taskCode: `POL-READ-${policy.Id}`,
            title: `Read ${policy.Title}`,
            description: `Review and read the ${policy.Title} document`,
            instructions: policy.Description || 'Access the policy through the Policy Hub and read thoroughly',
            category: 'Policy Compliance',
            priority: policy.IsMandatory ? 'High' : 'Normal',
            dueDate: this.calculateDueDate(startDate, policy.ReadingDueDays || 7),
            policyId: policy.Id,
            policyTitle: policy.Title,
            taskType: 'Read',
            estimatedHours: 0.5
          }
        );
        taskIds.push(readTaskId);

        // 2. Create "Acknowledge Policy" task (depends on reading)
        if (policy.RequiresAcknowledgement) {
          const ackTaskId = await this.createPolicyTask(
            processId,
            assignedToId,
            employeeEmail,
            {
              taskCode: `POL-ACK-${policy.Id}`,
              title: `Acknowledge ${policy.Title}`,
              description: `Formally acknowledge understanding of ${policy.Title}`,
              instructions: 'Click the Acknowledge button after reading the policy',
              category: 'Policy Compliance',
              priority: 'Critical',
              dueDate: this.calculateDueDate(startDate, policy.AcknowledgementDueDays || 10),
              policyId: policy.Id,
              policyTitle: policy.Title,
              taskType: 'Acknowledge',
              dependsOn: `POL-READ-${policy.Id}`,
              estimatedHours: 0.1
            }
          );
          taskIds.push(ackTaskId);
        }

        // 3. Create "Complete Quiz" task (depends on acknowledgement)
        if (policy.HasQuiz) {
          const quizTaskId = await this.createPolicyTask(
            processId,
            assignedToId,
            employeeEmail,
            {
              taskCode: `POL-QUIZ-${policy.Id}`,
              title: `Complete ${policy.Title} Quiz`,
              description: `Pass the ${policy.Title} assessment quiz`,
              instructions: `Score ${policy.PassingScore || 70}% or higher to complete. Retakes allowed.`,
              category: 'Policy Compliance',
              priority: 'High',
              dueDate: this.calculateDueDate(startDate, policy.QuizDueDays || 14),
              policyId: policy.Id,
              policyTitle: policy.Title,
              taskType: 'Quiz',
              dependsOn: policy.RequiresAcknowledgement ? `POL-ACK-${policy.Id}` : `POL-READ-${policy.Id}`,
              estimatedHours: 0.5
            }
          );
          taskIds.push(quizTaskId);
        }
      }

      return taskIds;
    } catch (error) {
      console.error('Error creating policy tasks for process:', error);
      throw error;
    }
  }

  /**
   * Create a single policy task assignment
   */
  private async createPolicyTask(
    processId: number,
    assignedToId: number,
    assignedToEmail: string,
    taskDetails: IPolicyTaskDetails
  ): Promise<number> {
    try {
      const result = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.add({
          Title: taskDetails.title,
          ProcessIDId: processId,
          TaskCode: taskDetails.taskCode,
          Category: taskDetails.category,
          Description: taskDetails.description,
          Instructions: taskDetails.instructions,
          AssignedToId: assignedToId,
          AssignedDate: new Date().toISOString(),
          DueDate: taskDetails.dueDate.toISOString(),
          StartDate: new Date().toISOString(),
          Priority: taskDetails.priority,
          Status: 'Not Started',
          PercentComplete: 0,
          EstimatedHours: taskDetails.estimatedHours,
          PolicyId: taskDetails.policyId,
          PolicyTitle: taskDetails.policyTitle,
          PolicyTaskType: taskDetails.taskType,
          DependsOn: taskDetails.dependsOn || null,
          RequiresApproval: false,
          IsBlocked: !!taskDetails.dependsOn // Blocked if it has dependency
        });

      // Log activity
      await this.analyticsService.logActivity(
        assignedToId,
        'Task Assigned',
        taskDetails.policyId,
        taskDetails.policyTitle,
        0,
        assignedToEmail.split('@')[0] // Extract department from email if needed
      );

      return result.data.Id;
    } catch (error) {
      console.error('Error creating policy task:', error);
      throw error;
    }
  }

  /**
   * Get mandatory policies for a department/role
   */
  private async getMandatoryPoliciesForRole(
    department: string,
    jobTitle: string
  ): Promise<IMandatoryPolicy[]> {
    try {
      // Query JML_Policies list for mandatory policies
      // Filter by target audience (department/role)
      const policies = await this.sp.web.lists
        .getByTitle('JML_Policies')
        .items.select(
          'Id',
          'Title',
          'Description',
          'IsMandatory',
          'TargetAudience',
          'TargetDepartments',
          'RequiresAcknowledgement',
          'HasQuiz',
          'PassingScore',
          'ReadingDueDays',
          'AcknowledgementDueDays',
          'QuizDueDays',
          'IsActive'
        )
        .filter(`IsActive eq true and IsMandatory eq true`)
        .orderBy('Priority', false)();

      // Filter by department (if TargetDepartments contains the department)
      const filtered = policies.filter((policy: any) => {
        if (!policy.TargetDepartments) return true; // If no target, applies to all
        const departments = policy.TargetDepartments.split(';');
        return departments.includes(department) || departments.includes('All');
      });

      return filtered as IMandatoryPolicy[];
    } catch (error) {
      console.error('Error getting mandatory policies:', error);
      // Return empty array if list doesn't exist yet
      return [];
    }
  }

  /**
   * Calculate due date from start date + days
   */
  private calculateDueDate(startDate: Date, daysToAdd: number): Date {
    const dueDate = new Date(startDate);
    dueDate.setDate(dueDate.getDate() + daysToAdd);
    return dueDate;
  }

  /**
   * Update policy task status when policy activity occurs
   */
  public async syncPolicyActivityToTask(
    processId: number,
    policyId: number,
    activityType: 'Read' | 'Acknowledge' | 'QuizPassed',
    userId: number,
    quizScore?: number
  ): Promise<void> {
    try {
      // Find the corresponding task assignment
      const taskCode =
        activityType === 'Read'
          ? `POL-READ-${policyId}`
          : activityType === 'Acknowledge'
          ? `POL-ACK-${policyId}`
          : `POL-QUIZ-${policyId}`;

      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(
          `ProcessIDId eq ${processId} and TaskCode eq '${taskCode}' and AssignedToId eq ${userId}`
        )
        .top(1)();

      if (tasks.length === 0) {
        console.warn(`No task found for ${taskCode} in process ${processId}`);
        return;
      }

      const task = tasks[0];

      // Update task status
      const updates: any = {
        Status: 'Completed',
        PercentComplete: 100,
        ActualCompletionDate: new Date().toISOString()
      };

      if (activityType === 'QuizPassed' && quizScore !== undefined) {
        updates.CompletionNotes = `Quiz passed with score: ${quizScore}%`;
      }

      await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.getById(task.Id)
        .update(updates);

      // Unblock dependent tasks
      await this.unblockDependentTasks(processId, taskCode);

      // Update process progress
      await this.updateProcessProgress(processId);
    } catch (error) {
      console.error('Error syncing policy activity to task:', error);
      throw error;
    }
  }

  /**
   * Unblock tasks that depend on the completed task
   */
  private async unblockDependentTasks(
    processId: number,
    completedTaskCode: string
  ): Promise<void> {
    try {
      const dependentTasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(
          `ProcessIDId eq ${processId} and DependsOn eq '${completedTaskCode}' and IsBlocked eq true`
        )();

      for (const task of dependentTasks) {
        await this.sp.web.lists
          .getByTitle('JML_TaskAssignments')
          .items.getById(task.Id)
          .update({
            IsBlocked: false,
            Status: 'Not Started'
          });
      }
    } catch (error) {
      console.error('Error unblocking dependent tasks:', error);
    }
  }

  /**
   * Update overall process progress after task completion
   */
  private async updateProcessProgress(processId: number): Promise<void> {
    try {
      // Get all tasks for the process
      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(`ProcessIDId eq ${processId}`)
        .select('Id', 'Status', 'PercentComplete')();

      const totalTasks = tasks.length;
      const completedTasks = tasks.filter((t: any) => t.Status === 'Completed').length;
      const progressPercentage = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;

      // Update process
      await this.sp.web.lists
        .getByTitle('JML_Processes')
        .items.getById(processId)
        .update({
          TotalTasks: totalTasks,
          CompletedTasks: completedTasks,
          ProgressPercentage: progressPercentage
        });
    } catch (error) {
      console.error('Error updating process progress:', error);
    }
  }

  /**
   * Get policy compliance status for an employee
   */
  public async getEmployeePolicyCompliance(
    employeeEmail: string,
    processId?: number
  ): Promise<IEmployeePolicyCompliance> {
    try {
      // Get user
      const user = await this.sp.web.siteUsers.getByEmail(employeeEmail)();

      // Get all policy tasks for this user
      let filter = `AssignedToId eq ${user.Id} and Category eq 'Policy Compliance'`;
      if (processId) {
        filter += ` and ProcessIDId eq ${processId}`;
      }

      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(filter)
        .select(
          'Id',
          'Title',
          'Status',
          'Priority',
          'DueDate',
          'PolicyId',
          'PolicyTitle',
          'PolicyTaskType',
          'PercentComplete',
          'IsBlocked'
        )
        .orderBy('DueDate', true)();

      const required = tasks.length;
      const completed = tasks.filter((t: any) => t.Status === 'Completed').length;
      const inProgress = tasks.filter((t: any) => t.Status === 'In Progress').length;
      const overdue = tasks.filter((t: any) => {
        if (t.Status === 'Completed') return false;
        return new Date(t.DueDate) < new Date();
      }).length;

      return {
        employeeEmail,
        userId: user.Id,
        requiredPolicies: required,
        completedPolicies: completed,
        inProgressPolicies: inProgress,
        overduePolicies: overdue,
        complianceRate: required > 0 ? Math.round((completed / required) * 100) : 100,
        tasks: tasks as IPolicyTaskStatus[]
      };
    } catch (error) {
      console.error('Error getting employee policy compliance:', error);
      throw error;
    }
  }

  /**
   * Get policy compliance for all users in a department
   */
  public async getDepartmentPolicyCompliance(
    department: string,
    managerId: number
  ): Promise<IDepartmentPolicyCompliance> {
    try {
      // Get all processes for the department
      const processes = await this.sp.web.lists
        .getByTitle('JML_Processes')
        .items.filter(`Department eq '${department}'`)
        .select('Id', 'EmployeeEmail', 'EmployeeName')();

      const teamMembers: ITeamMemberCompliance[] = [];

      for (const process of processes) {
        try {
          const compliance = await this.getEmployeePolicyCompliance(
            process.EmployeeEmail,
            process.Id
          );

          teamMembers.push({
            employeeName: process.EmployeeName,
            employeeEmail: process.EmployeeEmail,
            processId: process.Id,
            compliance
          });
        } catch (err) {
          console.warn(`Could not get compliance for ${process.EmployeeEmail}:`, err);
        }
      }

      // Calculate aggregates
      const totalMembers = teamMembers.length;
      const avgCompliance =
        totalMembers > 0
          ? Math.round(
              teamMembers.reduce((sum, m) => sum + m.compliance.complianceRate, 0) / totalMembers
            )
          : 0;
      const membersAtRisk = teamMembers.filter(
        (m) => m.compliance.overduePolicies > 0 || m.compliance.complianceRate < 70
      ).length;

      return {
        department,
        totalMembers,
        averageComplianceRate: avgCompliance,
        membersAtRisk,
        teamMembers
      };
    } catch (error) {
      console.error('Error getting department policy compliance:', error);
      throw error;
    }
  }

  /**
   * Get policy compliance status for a JML process
   */
  public async getPolicyComplianceStatus(
    processId: number
  ): Promise<{ totalPolicies: number; completedPolicies: number; complianceRate: number }> {
    try {
      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(`ProcessIDId eq ${processId} and Category eq 'Policy Compliance'`)
        .select('Id', 'Status')();

      const totalPolicies = tasks.length;
      const completedPolicies = tasks.filter((t: any) => t.Status === 'Completed').length;
      const complianceRate = totalPolicies > 0 ? Math.round((completedPolicies / totalPolicies) * 100) : 100;

      return { totalPolicies, completedPolicies, complianceRate };
    } catch (error) {
      console.error('Error getting policy compliance status:', error);
      return { totalPolicies: 0, completedPolicies: 0, complianceRate: 100 };
    }
  }

  /**
   * Get policy tasks formatted for display in OnboardingTracker
   */
  public async getPolicyTasksForDisplay(
    processId: number,
    userId: number
  ): Promise<IPolicyTaskDisplay[]> {
    try {
      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(
          `ProcessIDId eq ${processId} and AssignedToId eq ${userId} and Category eq 'Policy Compliance'`
        )
        .select(
          'Id',
          'Title',
          'Description',
          'Instructions',
          'Status',
          'Priority',
          'DueDate',
          'PolicyId',
          'PolicyTitle',
          'PolicyTaskType',
          'PercentComplete',
          'IsBlocked',
          'DependsOn',
          'CompletionNotes'
        )
        .orderBy('DueDate', true)();

      return tasks.map((task: any) => ({
        id: task.Id,
        title: task.Title,
        description: task.Description,
        instructions: task.Instructions,
        status: task.Status,
        priority: task.Priority,
        dueDate: new Date(task.DueDate),
        policyId: task.PolicyId,
        policyTitle: task.PolicyTitle,
        taskType: task.PolicyTaskType,
        percentComplete: task.PercentComplete,
        isBlocked: task.IsBlocked,
        dependsOn: task.DependsOn,
        completionNotes: task.CompletionNotes,
        isOverdue: new Date(task.DueDate) < new Date() && task.Status !== 'Completed',
        canComplete: !task.IsBlocked && task.Status !== 'Completed'
      }));
    } catch (error) {
      console.error('Error getting policy tasks for display:', error);
      return [];
    }
  }

  /**
   * Check for overdue policy tasks and create violations
   */
  public async checkAndCreateViolations(processId: number): Promise<number> {
    try {
      // Get overdue policy tasks
      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(
          `ProcessIDId eq ${processId} and Category eq 'Policy Compliance' and Status ne 'Completed'`
        )
        .select(
          'Id',
          'Title',
          'PolicyId',
          'PolicyTitle',
          'PolicyTaskType',
          'DueDate',
          'Priority',
          'AssignedTo/Id',
          'AssignedTo/Title',
          'AssignedTo/EMail'
        )
        .expand('AssignedTo')();

      let violationsCreated = 0;

      for (const task of tasks) {
        const dueDate = new Date(task.DueDate);
        const now = new Date();

        if (dueDate < now) {
          // Task is overdue
          const daysOverdue = Math.floor((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));

          // Check if violation already exists
          const existing = await this.sp.web.lists
            .getByTitle('JML_ComplianceViolations')
            .items.filter(
              `PolicyId eq ${task.PolicyId} and UserEmail eq '${task.AssignedTo.EMail}' and Status ne 'Resolved'`
            )
            .top(1)();

          if (existing.length === 0) {
            // Create new violation
            await this.sp.web.lists.getByTitle('JML_ComplianceViolations').items.add({
              Title: `${task.PolicyTaskType} - ${task.PolicyTitle}`,
              UserName: task.AssignedTo.Title,
              UserEmail: task.AssignedTo.EMail,
              ViolationType:
                task.PolicyTaskType === 'Read'
                  ? 'Overdue Reading'
                  : task.PolicyTaskType === 'Acknowledge'
                  ? 'Missing Acknowledgement'
                  : 'Quiz Not Attempted',
              PolicyId: task.PolicyId,
              PolicyTitle: task.PolicyTitle,
              Severity: this.calculateViolationSeverity(daysOverdue, task.Priority),
              DetectedDate: new Date().toISOString(),
              DueDate: dueDate.toISOString(),
              DaysOverdue: daysOverdue,
              Status: 'Open',
              Description: `Policy task overdue by ${daysOverdue} day(s): ${task.Title}`
            });

            violationsCreated++;
          }
        }
      }

      return violationsCreated;
    } catch (error) {
      console.error('Error checking and creating violations:', error);
      return 0;
    }
  }

  /**
   * Calculate violation severity based on days overdue and priority
   */
  private calculateViolationSeverity(daysOverdue: number, priority: string): string {
    if (priority === 'Critical') {
      return daysOverdue > 3 ? 'Critical' : 'High';
    }
    if (daysOverdue > 10) return 'Critical';
    if (daysOverdue > 5) return 'High';
    if (daysOverdue > 2) return 'Medium';
    return 'Low';
  }
}

// Interfaces
export interface IPolicyTaskDetails {
  taskCode: string;
  title: string;
  description: string;
  instructions: string;
  category: string;
  priority: string;
  dueDate: Date;
  policyId: number;
  policyTitle: string;
  taskType: 'Read' | 'Acknowledge' | 'Quiz';
  dependsOn?: string;
  estimatedHours: number;
}

export interface IMandatoryPolicy {
  Id: number;
  Title: string;
  Description: string;
  IsMandatory: boolean;
  TargetAudience: string;
  TargetDepartments: string;
  RequiresAcknowledgement: boolean;
  HasQuiz: boolean;
  PassingScore: number;
  ReadingDueDays: number;
  AcknowledgementDueDays: number;
  QuizDueDays: number;
  IsActive: boolean;
}

export interface IPolicyTaskStatus {
  Id: number;
  Title: string;
  Status: string;
  Priority: string;
  DueDate: string;
  PolicyId: number;
  PolicyTitle: string;
  PolicyTaskType: string;
  PercentComplete: number;
  IsBlocked: boolean;
}

export interface IEmployeePolicyCompliance {
  employeeEmail: string;
  userId: number;
  requiredPolicies: number;
  completedPolicies: number;
  inProgressPolicies: number;
  overduePolicies: number;
  complianceRate: number;
  tasks: IPolicyTaskStatus[];
}

export interface ITeamMemberCompliance {
  employeeName: string;
  employeeEmail: string;
  processId: number;
  compliance: IEmployeePolicyCompliance;
}

export interface IDepartmentPolicyCompliance {
  department: string;
  totalMembers: number;
  averageComplianceRate: number;
  membersAtRisk: number;
  teamMembers: ITeamMemberCompliance[];
}

export interface IPolicyTaskDisplay {
  id: number;
  title: string;
  description: string;
  instructions: string;
  status: string;
  priority: string;
  dueDate: Date;
  policyId: number;
  policyTitle: string;
  taskType: 'Read' | 'Acknowledge' | 'Quiz';
  percentComplete: number;
  isBlocked: boolean;
  dependsOn: string | null;
  completionNotes: string;
  isOverdue: boolean;
  canComplete: boolean;
}
